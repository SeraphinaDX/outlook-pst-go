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
	clientSig   byte // Client signature (PC=0xBC, TC=0x7C, BTH=0xB5)
	rootHID     util.HeapID
	allocations []heapAllocation
	nextAllocID uint16
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
		clientSig:   clientSig,
		nextAllocID: 0, // MakeHeapID converts the zero-based allocation index to a non-zero HID
		format:      format,
	}
}

// Allocate allocates space in the heap and returns the HID.
func (w *HeapWriter) Allocate(data []byte) (util.HeapID, error) {
	if len(data) > disk.HeapMaxAllocSize {
		return 0, fmt.Errorf("allocation too large: %d bytes (max %d)", len(data), disk.HeapMaxAllocSize)
	}

	// Create HID: bits 0-4 reserved, bits 5-15 block index, bits 16-31 alloc index
	// For simple single-block heaps, block index is always 0
	hid := util.MakeHeapID(0, w.nextAllocID)
	w.nextAllocID++

	w.allocations = append(w.allocations, heapAllocation{
		hid:  hid,
		data: data,
	})

	return hid, nil
}

// SetRoot sets the root HID for the heap.
func (w *HeapWriter) SetRoot(hid util.HeapID) {
	w.rootHID = hid
}

// Build builds the heap data block.
// Returns the complete heap block data ready for writing.
func (w *HeapWriter) Build() ([]byte, error) {
	// Calculate total size needed
	headerSize := 12                            // HNHDR size
	pageMapSize := 4 + (len(w.allocations)+1)*2 // cAlloc + cFree + allocation offsets/end marker

	// Calculate data size
	totalDataSize := 0
	for _, alloc := range w.allocations {
		totalDataSize += len(alloc.data)
	}

	// Total size (excluding page map at end)
	totalSize := headerSize + totalDataSize + pageMapSize

	// Check if it fits in a single block
	maxSize := disk.MaxDataBlockSizeUnicode
	if w.format == disk.FormatANSI {
		maxSize = disk.MaxDataBlockSizeANSI
	}
	if totalSize > maxSize {
		return nil, fmt.Errorf("heap too large for single block: %d bytes", totalSize)
	}

	// Allocate buffer
	buf := make([]byte, totalSize)

	// Write HNHDR (heap header) - See [MS-PST] Section 2.3.1.2
	// ibHnpm: 0-2 (2 bytes) - Offset to page map
	// bSig: 2 (1 byte) - Heap signature (0xEC)
	// bClientSig: 3 (1 byte) - Client signature
	// hidUserRoot: 4-8 (4 bytes) - Root HID
	// rgbFillLevel: 8-12 (4 bytes) - Fill levels

	pageMapOffset := headerSize + totalDataSize
	binary.LittleEndian.PutUint16(buf[0:2], uint16(pageMapOffset))
	buf[2] = disk.HeapSignature // 0xEC
	buf[3] = w.clientSig
	binary.LittleEndian.PutUint32(buf[4:8], uint32(w.rootHID))
	// Fill levels left as zero (block 0 will be partially filled)

	// Write allocations and build page map
	currentOffset := headerSize
	offsets := make([]uint16, len(w.allocations)+1)

	for i, alloc := range w.allocations {
		offsets[i] = uint16(currentOffset)
		copy(buf[currentOffset:], alloc.data)
		currentOffset += len(alloc.data)
	}
	offsets[len(w.allocations)] = uint16(currentOffset) // End marker

	// Write page map: cAlloc, cFree, then rgibAlloc[cAlloc+1].
	binary.LittleEndian.PutUint16(buf[pageMapOffset:pageMapOffset+2], uint16(len(w.allocations))) //nolint:gosec
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
