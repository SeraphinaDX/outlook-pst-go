package ltp

import (
	"encoding/binary"
	"fmt"
	"sort"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// BTHWriter builds a B-Tree on Heap structure.
// See [MS-PST] Section 2.3.2 for the BTH specification.
type BTHWriter struct {
	heap     *HeapWriter
	keySize  byte // Key size in bytes (2, 4, 8, or 16)
	dataSize byte // Data size in bytes
	entries  []bthEntry
	format   disk.PSTFormat
}

// bthEntry represents an entry in the BTH.
type bthEntry struct {
	key  []byte
	data []byte
}

// NewBTHWriter creates a new BTH writer.
// keySize: size of keys in bytes (typically 2 for PropID-based keys)
// dataSize: size of data per entry (typically 6 for property context entries)
func NewBTHWriter(heap *HeapWriter, keySize, dataSize byte, format disk.PSTFormat) *BTHWriter {
	return &BTHWriter{
		heap:     heap,
		keySize:  keySize,
		dataSize: dataSize,
		format:   format,
	}
}

// Insert adds an entry to the BTH.
func (w *BTHWriter) Insert(key, data []byte) error {
	if len(key) != int(w.keySize) {
		return fmt.Errorf("key size mismatch: got %d, expected %d", len(key), w.keySize)
	}
	if len(data) != int(w.dataSize) {
		return fmt.Errorf("data size mismatch: got %d, expected %d", len(data), w.dataSize)
	}

	// Check for duplicate key
	for i, entry := range w.entries {
		if compareBytes(entry.key, key) == 0 {
			// Update existing entry
			w.entries[i].data = data
			return nil
		}
	}

	// Add new entry
	keyCopy := make([]byte, len(key))
	dataCopy := make([]byte, len(data))
	copy(keyCopy, key)
	copy(dataCopy, data)

	w.entries = append(w.entries, bthEntry{
		key:  keyCopy,
		data: dataCopy,
	})

	return nil
}

// InsertUint16Key inserts an entry with a uint16 key.
func (w *BTHWriter) InsertUint16Key(key uint16, data []byte) error {
	keyBytes := make([]byte, 2)
	binary.LittleEndian.PutUint16(keyBytes, key)
	return w.Insert(keyBytes, data)
}

// InsertUint32Key inserts an entry with a uint32 key.
func (w *BTHWriter) InsertUint32Key(key uint32, data []byte) error {
	keyBytes := make([]byte, 4)
	binary.LittleEndian.PutUint32(keyBytes, key)
	return w.Insert(keyBytes, data)
}

// Delete removes an entry from the BTH.
func (w *BTHWriter) Delete(key []byte) error {
	for i, entry := range w.entries {
		if compareBytes(entry.key, key) == 0 {
			w.entries = append(w.entries[:i], w.entries[i+1:]...)
			return nil
		}
	}
	return fmt.Errorf("key not found")
}

// DeleteUint16Key removes an entry by uint16 key.
func (w *BTHWriter) DeleteUint16Key(key uint16) error {
	keyBytes := make([]byte, 2)
	binary.LittleEndian.PutUint16(keyBytes, key)
	return w.Delete(keyBytes)
}

// Build builds the BTH structure and returns the root HID.
func (w *BTHWriter) Build() (util.HeapID, error) {
	if len(w.entries) == 0 {
		// Create empty BTH header
		return w.createEmptyBTH()
	}

	// Sort entries by key
	sort.Slice(w.entries, func(i, j int) bool {
		return compareBytes(w.entries[i].key, w.entries[j].key) < 0
	})

	// For small BTH, create single-level structure
	// Calculate max entries per leaf block
	entrySize := int(w.keySize + w.dataSize)
	maxEntriesPerBlock := disk.HeapMaxAllocSize / entrySize

	if len(w.entries) <= maxEntriesPerBlock {
		return w.buildSingleLevel()
	}

	// Multi-level BTH (more complex, implement as needed)
	return w.buildMultiLevel(maxEntriesPerBlock)
}

// createEmptyBTH creates an empty BTH with just a header.
func (w *BTHWriter) createEmptyBTH() (util.HeapID, error) {
	// BTH header (BTHHEADER) - 8 bytes
	// See [MS-PST] Section 2.3.2.1
	header := make([]byte, 8)
	header[0] = 0xB5       // bType - BTH signature
	header[1] = w.keySize  // cbKey
	header[2] = w.dataSize // cbEnt (data size)
	header[3] = 0          // bIdxLevels (0 = leaf only)
	// hidRoot: 4-8 (4 bytes) - 0 for empty BTH
	binary.LittleEndian.PutUint32(header[4:8], 0)

	hid, err := w.heap.Allocate(header)
	if err != nil {
		return 0, err
	}

	w.heap.SetRoot(hid)
	return hid, nil
}

// buildSingleLevel builds a BTH with only leaf entries.
func (w *BTHWriter) buildSingleLevel() (util.HeapID, error) {
	entrySize := int(w.keySize + w.dataSize)

	// Build leaf data
	leafData := make([]byte, len(w.entries)*entrySize)
	for i, entry := range w.entries {
		offset := i * entrySize
		copy(leafData[offset:], entry.key)
		copy(leafData[offset+int(w.keySize):], entry.data)
	}

	// Allocate leaf block
	leafHID, err := w.heap.Allocate(leafData)
	if err != nil {
		return 0, err
	}

	// Build BTH header pointing to leaf
	header := make([]byte, 8)
	header[0] = 0xB5       // bType
	header[1] = w.keySize  // cbKey
	header[2] = w.dataSize // cbEnt
	header[3] = 0          // bIdxLevels (0 = leaf at root)
	binary.LittleEndian.PutUint32(header[4:8], uint32(leafHID))

	headerHID, err := w.heap.Allocate(header)
	if err != nil {
		return 0, err
	}

	w.heap.SetRoot(headerHID)
	return headerHID, nil
}

// buildMultiLevel builds a multi-level BTH for many entries.
func (w *BTHWriter) buildMultiLevel(maxPerBlock int) (util.HeapID, error) {
	entrySize := int(w.keySize + w.dataSize)

	// Split entries into leaf blocks
	var leafHIDs []util.HeapID
	var leafMaxKeys [][]byte

	for i := 0; i < len(w.entries); i += maxPerBlock {
		end := i + maxPerBlock
		if end > len(w.entries) {
			end = len(w.entries)
		}
		blockEntries := w.entries[i:end]

		// Build leaf data
		leafData := make([]byte, len(blockEntries)*entrySize)
		for j, entry := range blockEntries {
			offset := j * entrySize
			copy(leafData[offset:], entry.key)
			copy(leafData[offset+int(w.keySize):], entry.data)
		}

		leafHID, err := w.heap.Allocate(leafData)
		if err != nil {
			return 0, err
		}

		leafHIDs = append(leafHIDs, leafHID)
		leafMaxKeys = append(leafMaxKeys, blockEntries[len(blockEntries)-1].key)
	}

	// Build intermediate/root level
	return w.buildIntermediateLevel(leafHIDs, leafMaxKeys, 1)
}

// buildIntermediateLevel builds intermediate BTH levels.
func (w *BTHWriter) buildIntermediateLevel(childHIDs []util.HeapID, maxKeys [][]byte, level int) (util.HeapID, error) {
	if len(childHIDs) == 1 {
		// Single child - build header pointing to it
		header := make([]byte, 8)
		header[0] = 0xB5
		header[1] = w.keySize
		header[2] = w.dataSize
		// level counts the recursion step that produced this child. The child
		// itself is the root index node, so bIdxLevels is one less. For the
		// common split case (leaf blocks -> one index block), level is 2 here
		// and the on-disk BTH must advertise exactly one intermediate level.
		header[3] = byte(level - 1) //nolint:gosec // BTH depth is bounded and level >= 1
		binary.LittleEndian.PutUint32(header[4:8], uint32(childHIDs[0]))

		headerHID, err := w.heap.Allocate(header)
		if err != nil {
			return 0, err
		}

		w.heap.SetRoot(headerHID)
		return headerHID, nil
	}

	// Build intermediate entries (key + HID)
	// Intermediate entry: key(keySize) + HID(4)
	intEntrySize := int(w.keySize) + 4
	maxPerBlock := disk.HeapMaxAllocSize / intEntrySize

	var parentHIDs []util.HeapID
	var parentMaxKeys [][]byte

	for i := 0; i < len(childHIDs); i += maxPerBlock {
		end := i + maxPerBlock
		if end > len(childHIDs) {
			end = len(childHIDs)
		}

		// Build intermediate block
		blockData := make([]byte, (end-i)*intEntrySize)
		for j := i; j < end; j++ {
			offset := (j - i) * intEntrySize
			copy(blockData[offset:], maxKeys[j])
			binary.LittleEndian.PutUint32(blockData[offset+int(w.keySize):], uint32(childHIDs[j]))
		}

		blockHID, err := w.heap.Allocate(blockData)
		if err != nil {
			return 0, err
		}

		parentHIDs = append(parentHIDs, blockHID)
		parentMaxKeys = append(parentMaxKeys, maxKeys[end-1])
	}

	return w.buildIntermediateLevel(parentHIDs, parentMaxKeys, level+1)
}

// Count returns the number of entries.
func (w *BTHWriter) Count() int {
	return len(w.entries)
}

// compareBytes compares two byte slices.
func compareBytes(a, b []byte) int {
	minLen := len(a)
	if len(b) < minLen {
		minLen = len(b)
	}
	for i := 0; i < minLen; i++ {
		if a[i] < b[i] {
			return -1
		}
		if a[i] > b[i] {
			return 1
		}
	}
	if len(a) < len(b) {
		return -1
	}
	if len(a) > len(b) {
		return 1
	}
	return 0
}

// CreatePropertyContextBTH creates a BTH configured for Property Context.
// Keys are 2-byte PropIDs, data is 6-byte property records.
func CreatePropertyContextBTH(heap *HeapWriter, format disk.PSTFormat) *BTHWriter {
	return NewBTHWriter(heap, 2, 6, format)
}

// CreateRowIndexBTH creates a BTH for table row indexing.
// Keys are 4-byte row IDs, data is 4-byte row indices.
func CreateRowIndexBTH(heap *HeapWriter, format disk.PSTFormat) *BTHWriter {
	return NewBTHWriter(heap, 4, 4, format)
}
