package ltp

import (
	"encoding/binary"
	"fmt"
	"iter"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// ColumnDescriptor describes a column in a table.
// See [MS-PST] Section 2.3.4.2 - TCOLDESC structure.
type ColumnDescriptor struct {
	PropType  PropType // wPropType - Property type from [MS-OXCDATA]
	PropID    PropID   // wPropId - Property ID
	Offset    uint16   // ibData - Offset in row data
	Size      byte     // cbData - Column width in bytes
	BitOffset byte     // iBit - Bit position in cell existence bitmap (CEB)
}

// TCHeader represents the Table Context header.
// See [MS-PST] Section 2.3.4.1 - TCINFO structure.
type TCHeader struct {
	Signature     byte               // bType - Must be 0x7C (TC signature)
	NumColumns    byte               // cCols - Number of columns
	SizeOffsets   [4]uint16          // rgib - Offsets: 4-byte cols | 2-byte | 1-byte | CEB
	RowBTHID      util.HeapID        // hidRowIndex - HID of row index BTH
	RowMatrixHNID util.HeapNodeID    // hnidRowData - HNID of row data (heap or subnode)
	Columns       []ColumnDescriptor // rgTCOLDESC - Array of column descriptors
}

// Table represents a Table Context (TC).
// See [MS-PST] Section 2.3.4 - Table Context (TC).
// TC is a complex data structure used to store tabular data such as
// folder contents, recipient tables, and attachment tables.
type Table struct {
	heap   *HeapOnNode
	node   *ndb.Node
	header *TCHeader
	rowBTH *BTH // Row index BTH - maps dwRowID to row index

	// Row data - RowMatrix. See [MS-PST] Section 2.3.4.4.
	// May be stored in heap allocation or subnode depending on size.
	rowMatrix     []byte
	rowMatrixNode *ndb.Node // If row matrix is in a subnode
}

// TableRow represents a row in the table.
// See [MS-PST] Section 2.3.4.4 - RowMatrix for row data layout.
type TableRow struct {
	table   *Table
	rowID   uint32 // dwRowID - unique row identifier
	rowData []byte // Row data including CEB (Cell Existence Bitmap)
}

// NewTable creates a Table from a node.
func NewTable(node *ndb.Node) (*Table, error) {
	heap, err := NewHeapOnNode(node)
	if err != nil {
		return nil, fmt.Errorf("failed to create heap: %w", err)
	}

	// Verify this is a Table Context
	clientSig := heap.ClientSignature()
	if clientSig != disk.HeapSigTC {
		return nil, fmt.Errorf("invalid TC signature: got 0x%02X, expected 0x%02X", clientSig, disk.HeapSigTC)
	}

	// Read TC header from heap root
	rootID := heap.RootID()
	headerData, err := heap.Read(rootID)
	if err != nil {
		return nil, fmt.Errorf("failed to read TC header: %w", err)
	}

	header, err := parseTCHeader(headerData)
	if err != nil {
		return nil, fmt.Errorf("failed to parse TC header: %w", err)
	}

	// Create row index BTH
	rowBTH, err := NewBTH(heap, header.RowBTHID)
	if err != nil {
		return nil, fmt.Errorf("failed to create row BTH: %w", err)
	}

	t := &Table{
		heap:   heap,
		node:   node,
		header: header,
		rowBTH: rowBTH,
	}

	// Load row matrix
	if err := t.loadRowMatrix(); err != nil {
		return nil, fmt.Errorf("failed to load row matrix: %w", err)
	}

	return t, nil
}

// parseTCHeader parses the Table Context header.
func parseTCHeader(data []byte) (*TCHeader, error) {
	if len(data) < 22 {
		return nil, fmt.Errorf("TC header too small: %d bytes", len(data))
	}

	h := &TCHeader{
		Signature:  data[0],
		NumColumns: data[1],
	}

	if h.Signature != disk.HeapSigTC {
		return nil, fmt.Errorf("invalid TC signature: 0x%02X", h.Signature)
	}

	// Parse size offsets
	for i := 0; i < 4; i++ {
		h.SizeOffsets[i] = binary.LittleEndian.Uint16(data[2+i*2 : 4+i*2])
	}

	h.RowBTHID = util.HeapID(binary.LittleEndian.Uint32(data[10:14]))
	h.RowMatrixHNID = util.HeapNodeID(binary.LittleEndian.Uint32(data[14:18]))

	// Skip 4 bytes of unused/reserved
	// Columns start at offset 22

	h.Columns = make([]ColumnDescriptor, h.NumColumns)
	colOffset := 22
	for i := 0; i < int(h.NumColumns); i++ {
		if colOffset+8 > len(data) {
			return nil, fmt.Errorf("TC header truncated at column %d", i)
		}
		h.Columns[i] = ColumnDescriptor{
			PropType:  PropType(binary.LittleEndian.Uint16(data[colOffset : colOffset+2])),
			PropID:    PropID(binary.LittleEndian.Uint16(data[colOffset+2 : colOffset+4])),
			Offset:    binary.LittleEndian.Uint16(data[colOffset+4 : colOffset+6]),
			Size:      data[colOffset+6],
			BitOffset: data[colOffset+7],
		}
		colOffset += 8
	}

	return h, nil
}

// loadRowMatrix loads the row data storage.
func (t *Table) loadRowMatrix() error {
	if t.header.RowMatrixHNID == 0 {
		// Empty table
		return nil
	}

	if t.header.RowMatrixHNID.IsHeapID() {
		// Row data is in the heap
		data, err := t.heap.Read(t.header.RowMatrixHNID.ToHeapID())
		if err != nil {
			return err
		}
		t.rowMatrix = data
	} else {
		// Row data is in a subnode
		nid := t.header.RowMatrixHNID.ToNodeID()
		subnode, err := t.node.LookupSubnode(nid)
		if err != nil {
			return fmt.Errorf("failed to find row matrix subnode 0x%X: %w", nid, err)
		}
		t.rowMatrixNode = subnode
	}

	return nil
}

// RowCount returns the number of rows in the table.
func (t *Table) RowCount() int {
	count := 0
	for _, err := range t.rowBTH.Entries() {
		if err != nil {
			continue
		}
		count++
	}
	return count
}

// ColumnCount returns the number of columns.
func (t *Table) ColumnCount() int {
	return int(t.header.NumColumns)
}

// Columns returns the column descriptors.
func (t *Table) Columns() []ColumnDescriptor {
	return t.header.Columns
}

// FindColumn finds a column by property ID.
func (t *Table) FindColumn(propID PropID) *ColumnDescriptor {
	for i := range t.header.Columns {
		if t.header.Columns[i].PropID == propID {
			return &t.header.Columns[i]
		}
	}
	return nil
}

// RowSize returns the size of each row in bytes.
func (t *Table) RowSize() int {
	// The last size offset is the total row size including bitmap
	return int(t.header.SizeOffsets[3])
}

// Rows returns an iterator over all rows.
func (t *Table) Rows() iter.Seq2[*TableRow, error] {
	return func(yield func(*TableRow, error) bool) {
		for entry, err := range t.rowBTH.Entries() {
			if err != nil {
				yield(nil, err)
				return
			}

			// Key is row ID (4 bytes), value is row index (4 bytes)
			if len(entry.Key) < 4 || len(entry.Value) < 4 {
				continue
			}

			rowID := binary.LittleEndian.Uint32(entry.Key)
			rowIndex := binary.LittleEndian.Uint32(entry.Value)

			row, err := t.getRow(rowID, rowIndex)
			if err != nil {
				yield(nil, err)
				return
			}

			if !yield(row, nil) {
				return
			}
		}
	}
}

// getRow retrieves a row by its index.
func (t *Table) getRow(rowID, rowIndex uint32) (*TableRow, error) {
	rowSize := t.RowSize()
	offset := int(rowIndex) * rowSize

	var rowData []byte
	var err error

	if t.rowMatrixNode != nil {
		// Read from subnode
		rowData, err = t.rowMatrixNode.Read(uint64(offset), uint64(rowSize)) //nolint:gosec // G115: offset from row index * row size, bounded
	} else if t.rowMatrix != nil {
		// Read from heap data
		if offset+rowSize > len(t.rowMatrix) {
			return nil, fmt.Errorf("row index out of bounds: %d", rowIndex)
		}
		rowData = t.rowMatrix[offset : offset+rowSize]
	} else {
		// Empty table
		return nil, fmt.Errorf("table has no row data")
	}

	if err != nil {
		return nil, fmt.Errorf("failed to read row data: %w", err)
	}

	return &TableRow{
		table:   t,
		rowID:   rowID,
		rowData: rowData,
	}, nil
}

// GetRow retrieves a row by its row ID.
func (t *Table) GetRow(rowID uint32) (*TableRow, error) {
	// Look up row in BTH
	key := make([]byte, 4)
	binary.LittleEndian.PutUint32(key, rowID)

	value, err := t.rowBTH.Lookup(key)
	if err != nil {
		return nil, fmt.Errorf("row not found: %d", rowID)
	}

	rowIndex := binary.LittleEndian.Uint32(value)
	return t.getRow(rowID, rowIndex)
}

// RowID returns the row ID.
func (r *TableRow) RowID() uint32 {
	return r.rowID
}

// HasProperty returns true if the property exists in this row.
// See [MS-PST] Section 2.3.4.4.1 for row data layout:
// - 4-byte columns: 0 to rgib[TCI_4b] (SizeOffsets[0])
// - 2-byte columns: rgib[TCI_4b] to rgib[TCI_2b] (SizeOffsets[1])
// - 1-byte columns: rgib[TCI_2b] to rgib[TCI_1b] (SizeOffsets[2])
// - CEB: rgib[TCI_1b] to rgib[TCI_bm] (SizeOffsets[3])
func (r *TableRow) HasProperty(propID PropID) bool {
	col := r.table.FindColumn(propID)
	if col == nil {
		return false
	}

	// CEB (Cell Existence Block) starts at SizeOffsets[2] (end of 1-byte columns)
	bitmapOffset := int(r.table.header.SizeOffsets[2])
	if bitmapOffset >= len(r.rowData) {
		return false
	}

	byteIndex := int(col.BitOffset) / 8
	bitIndex := uint(col.BitOffset) % 8

	if bitmapOffset+byteIndex >= len(r.rowData) {
		return false
	}

	// [MS-PST] numbers iBit from the most-significant bit of each CEB byte.
	return (r.rowData[bitmapOffset+byteIndex] & (1 << (7 - bitIndex))) != 0
}

// GetRaw reads raw bytes for a column.
func (r *TableRow) GetRaw(propID PropID) ([]byte, error) {
	col := r.table.FindColumn(propID)
	if col == nil {
		return nil, fmt.Errorf("column not found: 0x%04X", propID)
	}

	if !r.HasProperty(propID) {
		return nil, fmt.Errorf("property not present in row: 0x%04X", propID)
	}

	offset := int(col.Offset)
	size := int(col.Size)

	if offset+size > len(r.rowData) {
		return nil, fmt.Errorf("column data out of bounds")
	}

	return r.rowData[offset : offset+size], nil
}

// GetInt32 reads an int32 value.
func (r *TableRow) GetInt32(propID PropID) (int32, error) {
	data, err := r.GetRaw(propID)
	if err != nil {
		return 0, err
	}
	if len(data) < 4 {
		return 0, fmt.Errorf("insufficient data for int32")
	}
	return int32(binary.LittleEndian.Uint32(data)), nil //nolint:gosec // G115: binary reinterpretation, same bit width
}

// GetInt64 reads an int64 value.
func (r *TableRow) GetInt64(propID PropID) (int64, error) {
	data, err := r.GetRaw(propID)
	if err != nil {
		return 0, err
	}
	if len(data) < 8 {
		return 0, fmt.Errorf("insufficient data for int64")
	}
	return int64(binary.LittleEndian.Uint64(data)), nil //nolint:gosec // G115: binary format reinterpretation
}

// GetHNID reads an HNID value (for variable-length properties).
func (r *TableRow) GetHNID(propID PropID) (util.HeapNodeID, error) {
	data, err := r.GetRaw(propID)
	if err != nil {
		return 0, err
	}
	if len(data) < 4 {
		return 0, fmt.Errorf("insufficient data for HNID")
	}
	return util.HeapNodeID(binary.LittleEndian.Uint32(data)), nil
}

// GetString reads a string value.
func (r *TableRow) GetString(propID PropID) (string, error) {
	col := r.table.FindColumn(propID)
	if col == nil {
		return "", fmt.Errorf("column not found: 0x%04X", propID)
	}

	hnid, err := r.GetHNID(propID)
	if err != nil {
		return "", err
	}

	if hnid == 0 {
		return "", nil
	}

	var data []byte
	if hnid.IsHeapID() {
		data, err = r.table.heap.Read(hnid.ToHeapID())
	} else {
		// Subnode
		nid := hnid.ToNodeID()
		subnode, err := r.table.node.LookupSubnode(nid)
		if err != nil {
			return "", fmt.Errorf("failed to find subnode 0x%X: %w", nid, err)
		}
		data, err = subnode.ReadAll()
		if err != nil {
			return "", err
		}
	}

	if err != nil {
		return "", err
	}

	// Decode based on property type
	if col.PropType == PropTypeString {
		return decodeUTF16LE(data), nil
	}
	return string(data), nil
}

// GetBinary reads binary data.
func (r *TableRow) GetBinary(propID PropID) ([]byte, error) {
	hnid, err := r.GetHNID(propID)
	if err != nil {
		return nil, err
	}

	if hnid == 0 {
		return nil, nil
	}

	if hnid.IsHeapID() {
		return r.table.heap.Read(hnid.ToHeapID())
	}

	// Subnode
	nid := hnid.ToNodeID()
	subnode, err := r.table.node.LookupSubnode(nid)
	if err != nil {
		return nil, fmt.Errorf("failed to find subnode 0x%X: %w", nid, err)
	}
	return subnode.ReadAll()
}
