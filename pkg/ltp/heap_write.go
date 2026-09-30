package ltp

import (
	"encoding/binary"
	"fmt"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// HeapWriter manages heap allocations within a node.
// It builds a Heap-on-Node (HN) structure for property storage.
type HeapWriter struct {
	clientSig    byte // Client signature (PC=0xBC, TC=0x7C, BTH=0xB5)
	rootHID      util.HeapID
	allocations  []heapAllocation
	currentBlock uint16
	format       disk.PSTFormat
}

// heapAllocation represents a single heap allocation.
type heapAllocation struct {
	hid        util.HeapID
	blockIndex uint16
	data       []byte
}

// NewHeapWriter creates a new heap writer.
func NewHeapWriter(clientSig byte, format disk.PSTFormat) *HeapWriter {
	return &HeapWriter{
		clientSig: clientSig,
		format:    format,
	}
}

// Allocate allocates space in the heap and returns the HID.
func (w *HeapWriter) Allocate(data []byte) (util.HeapID, error) {
	if len(data) > disk.HeapMaxAllocSize {
		return 0, fmt.Errorf("allocation too large: %d bytes (max %d)", len(data), disk.HeapMaxAllocSize)
	}

	block := w.currentBlock
	if !w.allocationFits(block, len(data)) {
		block++
		for !w.allocationFits(block, len(data)) {
			block++
		}
		w.currentBlock = block
	}

	allocIndex := uint16(0)
	for _, alloc := range w.allocations {
		if alloc.blockIndex == block {
			allocIndex++
		}
	}
	if allocIndex >= 0x7FF {
		return 0, fmt.Errorf("too many heap allocations in block %d", block)
	}

	hid := util.MakeHeapID(block, allocIndex)
	w.allocations = append(w.allocations, heapAllocation{
		hid:        hid,
		blockIndex: block,
		data:       append([]byte(nil), data...),
	})
	return hid, nil
}

func (w *HeapWriter) allocationFits(block uint16, dataLen int) bool {
	headerSize := heapPageHeaderSize(block)
	dataSize := dataLen
	count := 1
	for _, alloc := range w.allocations {
		if alloc.blockIndex == block {
			dataSize += len(alloc.data)
			count++
		}
	}

	pageMapOffset := headerSize + dataSize
	if pageMapOffset%2 != 0 {
		pageMapOffset++
	}
	pageMapSize := 4 + (count+1)*2
	return pageMapOffset+pageMapSize <= w.maxBlockSize()
}

func (w *HeapWriter) maxBlockSize() int {
	if w.format == disk.FormatANSI {
		return disk.MaxDataBlockSizeANSI
	}
	return disk.MaxDataBlockSizeUnicode
}

func heapPageHeaderSize(block uint16) int {
	switch {
	case block == 0:
		return 12
	case block >= 8 && (block-8)%128 == 0:
		return 66
	default:
		return 2
	}
}

// SetRoot sets the root HID for the heap.
func (w *HeapWriter) SetRoot(hid util.HeapID) {
	w.rootHID = hid
}

// Build builds the Heap-on-Node data stream. When the heap spans more than
// one NDB data block, each HN page is padded to the NDB block payload size so
// WriteExtendedBlockData preserves the page boundaries.
func (w *HeapWriter) Build() ([]byte, error) {
	maxBlock := uint16(0)
	for _, alloc := range w.allocations {
		if alloc.blockIndex > maxBlock {
			maxBlock = alloc.blockIndex
		}
	}

	pages := make([][]byte, int(maxBlock)+1)
	used := make([]int, len(pages))
	for block := uint16(0); block <= maxBlock; block++ {
		var allocs []heapAllocation
		for _, alloc := range w.allocations {
			if alloc.blockIndex == block {
				allocs = append(allocs, alloc)
			}
		}
		page, pageUsed, err := w.buildPage(block, allocs, block < maxBlock)
		if err != nil {
			return nil, err
		}
		pages[block] = page
		used[block] = pageUsed
	}

	w.writeFillLevels(pages, used)

	total := 0
	for _, page := range pages {
		total += len(page)
	}
	out := make([]byte, 0, total)
	for _, page := range pages {
		out = append(out, page...)
	}
	return out, nil
}

func (w *HeapWriter) buildPage(block uint16, allocs []heapAllocation, padToBlock bool) ([]byte, int, error) {
	headerSize := heapPageHeaderSize(block)
	offsets := make([]uint16, len(allocs)+1)
	current := headerSize
	for i, alloc := range allocs {
		offsets[i] = uint16(current) //nolint:gosec // bounded by NDB block size
		current += len(alloc.data)
	}
	offsets[len(allocs)] = uint16(current) //nolint:gosec

	pageMapOffset := current
	if pageMapOffset%2 != 0 {
		pageMapOffset++
	}
	pageMapSize := 4 + len(offsets)*2
	used := pageMapOffset + pageMapSize
	if used > w.maxBlockSize() {
		return nil, 0, fmt.Errorf("heap page %d too large: %d bytes", block, used)
	}

	size := used
	if padToBlock {
		size = w.maxBlockSize()
	}
	buf := make([]byte, size)

	binary.LittleEndian.PutUint16(buf[0:2], uint16(pageMapOffset)) //nolint:gosec
	switch {
	case block == 0:
		buf[2] = disk.HeapSignature
		buf[3] = w.clientSig
		binary.LittleEndian.PutUint32(buf[4:8], uint32(w.rootHID))
	case block >= 8 && (block-8)%128 == 0:
		// HNBITMAPHDR: ibHnpm followed by a 64-byte fill-level map.
	default:
		// HNPAGEHDR contains only ibHnpm.
	}

	current = headerSize
	for _, alloc := range allocs {
		copy(buf[current:], alloc.data)
		current += len(alloc.data)
	}

	binary.LittleEndian.PutUint16(buf[pageMapOffset:pageMapOffset+2], uint16(len(allocs))) //nolint:gosec
	binary.LittleEndian.PutUint16(buf[pageMapOffset+2:pageMapOffset+4], 0)
	for i, off := range offsets {
		binary.LittleEndian.PutUint16(buf[pageMapOffset+4+i*2:], off)
	}
	return buf, used, nil
}

func (w *HeapWriter) writeFillLevels(pages [][]byte, used []int) {
	for block := range pages {
		level := heapFillLevel(w.maxBlockSize() - used[block])
		if block < 8 {
			setFillNibble(pages[0][8:12], block, level)
			continue
		}
		mapStart := 8 + ((block - 8) / 128 * 128)
		if mapStart >= len(pages) {
			continue
		}
		index := block - mapStart
		setFillNibble(pages[mapStart][2:66], index, level)
	}
}

func setFillNibble(buf []byte, index int, level byte) {
	if index < 0 || index/2 >= len(buf) {
		return
	}
	if index%2 == 0 {
		buf[index/2] = (buf[index/2] & 0xF0) | (level & 0x0F)
	} else {
		buf[index/2] = (buf[index/2] & 0x0F) | ((level & 0x0F) << 4)
	}
}

func heapFillLevel(free int) byte {
	switch {
	case free >= 3584:
		return 0x0
	case free >= 2560:
		return 0x1
	case free >= 2048:
		return 0x2
	case free >= 1792:
		return 0x3
	case free >= 1536:
		return 0x4
	case free >= 1280:
		return 0x5
	case free >= 1024:
		return 0x6
	case free >= 768:
		return 0x7
	case free >= 512:
		return 0x8
	case free >= 256:
		return 0x9
	case free >= 128:
		return 0xA
	case free >= 64:
		return 0xB
	case free >= 32:
		return 0xC
	case free >= 16:
		return 0xD
	case free >= 8:
		return 0xE
	default:
		return 0xF
	}
}

// HeapBlockWriter writes multi-block heaps for large data.
type HeapBlockWriter struct {
	clientSig byte
	blocks    [][]byte
	format    disk.PSTFormat
}

// NewHeapBlockWriter creates a writer for multi-block heaps.
func NewHeapBlockWriter(clientSig byte, format disk.PSTFormat) *HeapBlockWriter {
	return &HeapBlockWriter{
		clientSig: clientSig,
		format:    format,
	}
}

// AddBlock adds a heap block.
func (w *HeapBlockWriter) AddBlock(data []byte) int {
	w.blocks = append(w.blocks, data)
	return len(w.blocks) - 1
}

// BuildBlocks returns all blocks ready for writing.
func (w *HeapBlockWriter) BuildBlocks() [][]byte {
	return w.blocks
}

// CreatePropertyContextHeap creates a heap configured for Property Context.
func CreatePropertyContextHeap(format disk.PSTFormat) *HeapWriter {
	return NewHeapWriter(disk.HeapSigPC, format)
}

// CreateTableContextHeap creates a heap configured for Table Context.
func CreateTableContextHeap(format disk.PSTFormat) *HeapWriter {
	return NewHeapWriter(disk.HeapSigTC, format)
}

// CreateBTHHeap creates a heap configured for B-Tree on Heap.
func CreateBTHHeap(format disk.PSTFormat) *HeapWriter {
	return NewHeapWriter(disk.HeapSigBTH, format)
}

// SerializeHeapID serializes a HeapID to bytes.
func SerializeHeapID(hid util.HeapID) []byte {
	buf := make([]byte, 4)
	binary.LittleEndian.PutUint32(buf, uint32(hid))
	return buf
}

// ParseHeapID parses a HeapID from bytes.
func ParseHeapID(data []byte) util.HeapID {
	if len(data) < 4 {
		return 0
	}
	return util.HeapID(binary.LittleEndian.Uint32(data))
}
