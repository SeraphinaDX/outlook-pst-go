package ltp

import (
	"encoding/binary"
	"fmt"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// HeapWriter manages heap allocations within a node.
// It builds a Heap-on-Node (HN) structure for property and table storage.
type HeapWriter struct {
	clientSig   byte
	rootHID     util.HeapID
	allocations []heapAllocation
	format      disk.PSTFormat
}

// heapAllocation represents a single heap allocation.
type heapAllocation struct {
	hid  util.HeapID
	data []byte
}

// NewHeapWriter creates a new heap writer.
func NewHeapWriter(clientSig byte, format disk.PSTFormat) *HeapWriter {
	return &HeapWriter{
		clientSig: clientSig,
		format:    format,
	}
}

// Allocate allocates space in the heap and returns the HID.
//
// Heap allocations are packed into HN pages. A HID identifies both the page
// and the allocation within that page, so once one page is full we can move
// to the next page without changing callers.
func (w *HeapWriter) Allocate(data []byte) (util.HeapID, error) {
	if len(data) > disk.HeapMaxAllocSize {
		return 0, fmt.Errorf("allocation too large: %d bytes (max %d)", len(data), disk.HeapMaxAllocSize)
	}

	pageIndex := uint16(0)
	allocIndex := uint16(0)
	if len(w.allocations) > 0 {
		last := w.allocations[len(w.allocations)-1].hid
		pageIndex = last.PageIndex()
		allocIndex = last.AllocIndex() + 1
	}

	if !w.fitsOnPage(pageIndex, len(data), int(allocIndex)+1) {
		if pageIndex == ^uint16(0) {
			return 0, fmt.Errorf("heap has too many pages")
		}
		pageIndex++
		allocIndex = 0
		if !w.fitsOnPage(pageIndex, len(data), 1) {
			return 0, fmt.Errorf("allocation does not fit on empty heap page: %d bytes", len(data))
		}
	}

	hid := util.MakeHeapID(pageIndex, allocIndex)
	w.allocations = append(w.allocations, heapAllocation{
		hid:  hid,
		data: append([]byte(nil), data...),
	})
	return hid, nil
}

func (w *HeapWriter) fitsOnPage(pageIndex uint16, newDataSize, newAllocCount int) bool {
	headerSize := 2
	if pageIndex == 0 {
		headerSize = 12
	}

	dataSize := 0
	for _, alloc := range w.allocations {
		if alloc.hid.PageIndex() == pageIndex {
			dataSize += len(alloc.data)
		}
	}

	// HNPAGEMAP = cAlloc(2) + cFree(2) + rgibAlloc[cAlloc+1].
	pageMapSize := 4 + (newAllocCount+1)*2
	total := headerSize + dataSize + newDataSize + pageMapSize
	return total <= w.maxPageSize()
}

func (w *HeapWriter) maxPageSize() int {
	if w.format == disk.FormatANSI {
		return disk.MaxDataBlockSizeANSI
	}
	return disk.MaxDataBlockSizeUnicode
}

// SetRoot sets the root HID for the heap.
func (w *HeapWriter) SetRoot(hid util.HeapID) {
	w.rootHID = hid
}

// Build serializes the complete Heap-on-Node.
//
// When the heap spans multiple pages, every page except the final page is
// padded to the PST data-block payload size. WriteExtendedBlockData therefore
// preserves HN page boundaries exactly, and HeapOnNode can later address them
// using the block index encoded in each HID.
func (w *HeapWriter) Build() ([]byte, error) {
	maxPage := uint16(0)
	if len(w.allocations) > 0 {
		maxPage = w.allocations[len(w.allocations)-1].hid.PageIndex()
	}

	var result []byte
	for pageIndex := uint16(0); ; pageIndex++ {
		pageAllocs := make([]heapAllocation, 0)
		for _, alloc := range w.allocations {
			if alloc.hid.PageIndex() == pageIndex {
				pageAllocs = append(pageAllocs, alloc)
			}
		}

		page, err := w.buildPage(pageIndex, pageAllocs, pageIndex < maxPage)
		if err != nil {
			return nil, err
		}
		result = append(result, page...)

		if pageIndex == maxPage {
			break
		}
	}
	return result, nil
}

func (w *HeapWriter) buildPage(pageIndex uint16, allocs []heapAllocation, pad bool) ([]byte, error) {
	headerSize := 2
	if pageIndex == 0 {
		headerSize = 12
	}

	totalDataSize := 0
	for _, alloc := range allocs {
		totalDataSize += len(alloc.data)
	}
	pageMapSize := 4 + (len(allocs)+1)*2
	usedSize := headerSize + totalDataSize + pageMapSize
	if usedSize > w.maxPageSize() {
		return nil, fmt.Errorf("heap page %d too large: %d bytes", pageIndex, usedSize)
	}

	pageSize := usedSize
	if pad {
		pageSize = w.maxPageSize()
	}
	buf := make([]byte, pageSize)

	pageMapOffset := headerSize + totalDataSize
	binary.LittleEndian.PutUint16(buf[0:2], uint16(pageMapOffset)) //nolint:gosec
	if pageIndex == 0 {
		buf[2] = disk.HeapSignature
		buf[3] = w.clientSig
		binary.LittleEndian.PutUint32(buf[4:8], uint32(w.rootHID))
	}

	currentOffset := headerSize
	offsets := make([]uint16, len(allocs)+1)
	for i, alloc := range allocs {
		if alloc.hid.AllocIndex() != uint16(i) { //nolint:gosec
			return nil, fmt.Errorf("non-contiguous heap allocation index on page %d", pageIndex)
		}
		offsets[i] = uint16(currentOffset) //nolint:gosec
		copy(buf[currentOffset:], alloc.data)
		currentOffset += len(alloc.data)
	}
	offsets[len(allocs)] = uint16(currentOffset) //nolint:gosec

	binary.LittleEndian.PutUint16(buf[pageMapOffset:pageMapOffset+2], uint16(len(allocs))) //nolint:gosec
	binary.LittleEndian.PutUint16(buf[pageMapOffset+2:pageMapOffset+4], 0)
	for i, off := range offsets {
		binary.LittleEndian.PutUint16(buf[pageMapOffset+4+i*2:], off)
	}

	return buf, nil
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
