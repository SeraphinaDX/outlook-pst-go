package ltp

import (
	"encoding/binary"
	"fmt"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// TableWriter builds a Table Context structure for writing.
// See [MS-PST] Section 2.3.4 for the Table Context specification.
type TableWriter struct {
	heap      *HeapWriter
	rowBTH    *BTHWriter
	format    disk.PSTFormat
	columns   []ColumnDef
	rows      []*tableRow
	nextRowID uint32
}

// ColumnDef defines a column in the table.
type ColumnDef struct {
	PropID    PropID
	PropType  PropType
	Size      int
	Offset    uint16
	BitOffset byte
}

type tableRow struct {
	rowID  uint32
	values map[PropID][]byte
}

// TableSubnode is a large Table Context value stored outside the HN.
type TableSubnode struct {
	NID  util.NodeID
	Data []byte
}

func NewTableWriter(format disk.PSTFormat) *TableWriter {
	heap := CreateTableContextHeap(format)
	return &TableWriter{
		heap:      heap,
		rowBTH:    CreateRowIndexBTH(heap, format),
		format:    format,
		nextRowID: 1,
	}
}

// NewTableWriterFromTable creates a writer populated with an existing table.
// Variable-sized values are copied by value rather than copying old HNIDs.
func NewTableWriterFromTable(table *Table, format disk.PSTFormat) (*TableWriter, error) {
	if table == nil {
		return nil, fmt.Errorf("table cannot be nil")
	}

	w := NewTableWriter(format)
	for _, col := range table.Columns() {
		w.AddColumn(col.PropID, col.PropType)
	}

	for row, err := range table.Rows() {
		if err != nil {
			return nil, fmt.Errorf("failed to read table row: %w", err)
		}

		rowID, err := w.AddRowWithID(row.RowID())
		if err != nil {
			return nil, err
		}

		for _, col := range table.Columns() {
			if !row.HasProperty(col.PropID) {
				continue
			}

			var data []byte
			if col.PropType.IsFixedSize() {
				data, err = row.GetRaw(col.PropID)
			} else {
				data, err = row.GetBinary(col.PropID)
			}
			if err != nil {
				return nil, fmt.Errorf("failed to copy property 0x%04X from row %d: %w", col.PropID, row.RowID(), err)
			}

			if err := w.SetRowValue(rowID, col.PropID, append([]byte(nil), data...)); err != nil {
				return nil, err
			}
		}
	}

	return w, nil
}

func (w *TableWriter) AddColumn(propID PropID, propType PropType) {
	size := propType.FixedSize()
	if size == 0 {
		size = 4
	}
	w.columns = append(w.columns, ColumnDef{PropID: propID, PropType: propType, Size: size})
}

func (w *TableWriter) AddRow() uint32 {
	for w.hasRow(w.nextRowID) {
		w.nextRowID++
	}
	rowID := w.nextRowID
	w.nextRowID++
	w.rows = append(w.rows, &tableRow{rowID: rowID, values: make(map[PropID][]byte)})
	return rowID
}

// AddRowWithID adds a row with a caller-supplied dwRowID.
// Hierarchy and contents tables use object NIDs as row IDs.
func (w *TableWriter) AddRowWithID(rowID uint32) (uint32, error) {
	if w.hasRow(rowID) {
		return 0, fmt.Errorf("row %d already exists", rowID)
	}

	w.rows = append(w.rows, &tableRow{rowID: rowID, values: make(map[PropID][]byte)})
	if rowID >= w.nextRowID && rowID != ^uint32(0) {
		w.nextRowID = rowID + 1
	}
	return rowID, nil
}

func (w *TableWriter) hasRow(rowID uint32) bool {
	for _, row := range w.rows {
		if row.rowID == rowID {
			return true
		}
	}
	return false
}

func (w *TableWriter) SetRowValue(rowID uint32, propID PropID, value []byte) error {
	for _, row := range w.rows {
		if row.rowID == rowID {
			row.values[propID] = value
			return nil
		}
	}
	return fmt.Errorf("row %d not found", rowID)
}

func (w *TableWriter) SetRowInt32(rowID uint32, propID PropID, value int32) error {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint32(data, uint32(value)) //nolint:gosec // bit-preserving conversion
	return w.SetRowValue(rowID, propID, data)
}

func (w *TableWriter) SetRowInt64(rowID uint32, propID PropID, value int64) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data, uint64(value)) //nolint:gosec // bit-preserving conversion
	return w.SetRowValue(rowID, propID, data)
}

func (w *TableWriter) SetRowBool(rowID uint32, propID PropID, value bool) error {
	data := make([]byte, 2)
	if value {
		binary.LittleEndian.PutUint16(data, 1)
	}
	return w.SetRowValue(rowID, propID, data)
}

func (w *TableWriter) SetRowTime(rowID uint32, propID PropID, value uint64) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data, value)
	return w.SetRowValue(rowID, propID, data)
}

func (w *TableWriter) SetRowString(rowID uint32, propID PropID, value string) error {
	return w.SetRowValue(rowID, propID, encodeUTF16LE(value))
}

func (w *TableWriter) SetRowBinary(rowID uint32, propID PropID, value []byte) error {
	return w.SetRowValue(rowID, propID, value)
}

func (w *TableWriter) DeleteRow(rowID uint32) {
	for i, row := range w.rows {
		if row.rowID == rowID {
			w.rows = append(w.rows[:i], w.rows[i+1:]...)
			return
		}
	}
}

func (w *TableWriter) Build() ([]byte, error) {
	data, _, err := w.build(false)
	return data, err
}

// BuildWithSubnodes builds a table while moving values that cannot fit in a
// single HN allocation into subnodes. The caller must attach the returned
// subnodes to the table node.
func (w *TableWriter) BuildWithSubnodes() ([]byte, []TableSubnode, error) {
	return w.build(true)
}

func (w *TableWriter) build(allowSubnodes bool) ([]byte, []TableSubnode, error) {
	if len(w.columns) == 0 {
		return nil, nil, fmt.Errorf("table must have at least one column")
	}
	if len(w.columns) > 255 {
		return nil, nil, fmt.Errorf("table has too many columns: %d", len(w.columns))
	}

	// Build from the logical rows each time. Folder tables are rewritten as
	// messages are added, and reusing the previous heap would retain stale
	// allocations and duplicate BTH structures.
	w.heap = CreateTableContextHeap(w.format)
	w.rowBTH = CreateRowIndexBTH(w.heap, w.format)

	end4, end2, end1 := w.calculateColumnOffsets()
	rowDataSize := int(end1)
	existenceBitmapSize := (len(w.columns) + 7) / 8
	rowSize := rowDataSize + existenceBitmapSize

	var subnodes []TableSubnode
	nextSubnodeIndex := uint32(2) // index 1 is reserved for an external RowMatrix
	rowMatrix := make([]byte, len(w.rows)*rowSize)

	for i, row := range w.rows {
		rowOffset := i * rowSize
		existenceBitmap := make([]byte, existenceBitmapSize)

		for _, col := range w.columns {
			value, exists := row.values[col.PropID]
			if !exists && col.PropID == PidTagLtpRowId {
				value = make([]byte, 4)
				binary.LittleEndian.PutUint32(value, row.rowID)
				exists = true
			}
			if !exists {
				continue
			}

			byteIdx := int(col.BitOffset) / 8
			bitIdx := uint(col.BitOffset) % 8
			existenceBitmap[byteIdx] |= 1 << (7 - bitIdx)

			valueOffset := rowOffset + int(col.Offset)
			if col.PropType.IsFixedSize() {
				copyLen := col.Size
				if len(value) < copyLen {
					copyLen = len(value)
				}
				copy(rowMatrix[valueOffset:valueOffset+col.Size], value[:copyLen])
				continue
			}

			var hnid util.HeapNodeID
			if len(value) > disk.HeapMaxAllocSize {
				if !allowSubnodes {
					return nil, nil, fmt.Errorf("row value too large for heap: %d bytes", len(value))
				}
				nid := util.MakeNID(util.NIDTypeLTP, nextSubnodeIndex)
				nextSubnodeIndex++
				subnodes = append(subnodes, TableSubnode{NID: nid, Data: append([]byte(nil), value...)})
				hnid = util.HeapNodeID(nid)
			} else {
				hid, err := w.heap.Allocate(value)
				if err != nil {
					return nil, nil, fmt.Errorf("failed to allocate row value: %w", err)
				}
				hnid = util.HeapNodeID(hid)
			}
			binary.LittleEndian.PutUint32(rowMatrix[valueOffset:valueOffset+4], uint32(hnid))
		}

		copy(rowMatrix[rowOffset+rowDataSize:rowOffset+rowSize], existenceBitmap)

		indexData := make([]byte, 4)
		binary.LittleEndian.PutUint32(indexData, uint32(i))
		if err := w.rowBTH.InsertUint32Key(row.rowID, indexData); err != nil {
			return nil, nil, fmt.Errorf("failed to insert row %d into BTH: %w", row.rowID, err)
		}
	}

	var rowMatrixHNID util.HeapNodeID
	if len(rowMatrix) > 0 {
		if len(rowMatrix) > disk.HeapMaxAllocSize {
			if !allowSubnodes {
				return nil, nil, fmt.Errorf("row matrix too large for heap: %d bytes", len(rowMatrix))
			}
			nid := util.MakeNID(util.NIDTypeLTP, 1)
			subnodes = append(subnodes, TableSubnode{NID: nid, Data: rowMatrix})
			rowMatrixHNID = util.HeapNodeID(nid)
		} else {
			hid, err := w.heap.Allocate(rowMatrix)
			if err != nil {
				return nil, nil, fmt.Errorf("failed to allocate row matrix: %w", err)
			}
			rowMatrixHNID = util.HeapNodeID(hid)
		}
	}

	rowBTHHID, err := w.rowBTH.Build()
	if err != nil {
		return nil, nil, fmt.Errorf("failed to build row BTH: %w", err)
	}

	tcInfo := w.buildTCInfo(rowBTHHID, rowMatrixHNID, end4, end2, end1, uint16(rowSize)) //nolint:gosec
	tcInfoHID, err := w.heap.Allocate(tcInfo)
	if err != nil {
		return nil, nil, fmt.Errorf("failed to allocate TCINFO: %w", err)
	}
	w.heap.SetRoot(tcInfoHID)

	data, err := w.heap.Build()
	if err != nil {
		return nil, nil, err
	}
	return data, subnodes, nil
}

// calculateColumnOffsets lays out 8/4-byte values (including HNIDs), then
// 2-byte values, then 1-byte values, followed by the cell-existence bitmap.
func (w *TableWriter) calculateColumnOffsets() (end4, end2, end1 uint16) {
	for i := range w.columns {
		w.columns[i].BitOffset = byte(i) //nolint:gosec // Build limits column count to 255
	}

	offset := uint16(0)
	assign := func(match func(int) bool) {
		for i := range w.columns {
			if !match(w.columns[i].Size) {
				continue
			}
			w.columns[i].Offset = offset
			offset += uint16(w.columns[i].Size) //nolint:gosec
		}
	}

	assign(func(size int) bool { return size >= 4 })
	end4 = offset
	assign(func(size int) bool { return size == 2 })
	end2 = offset
	assign(func(size int) bool { return size == 1 })
	end1 = offset
	return end4, end2, end1
}

func (w *TableWriter) buildColumnDescriptors() []byte {
	data := make([]byte, len(w.columns)*8)
	for i, col := range w.columns {
		offset := i * 8
		binary.LittleEndian.PutUint16(data[offset:offset+2], uint16(col.PropType))
		binary.LittleEndian.PutUint16(data[offset+2:offset+4], uint16(col.PropID))
		binary.LittleEndian.PutUint16(data[offset+4:offset+6], col.Offset)
		data[offset+6] = byte(col.Size) //nolint:gosec
		data[offset+7] = col.BitOffset
	}
	return data
}

// buildTCInfo keeps TCOLDESCs inline with TCINFO as required by [MS-PST].
func (w *TableWriter) buildTCInfo(rowBTHHID util.HeapID, rowMatrixHNID util.HeapNodeID, end4, end2, end1, rowSize uint16) []byte {
	data := make([]byte, 22+len(w.columns)*8)
	data[0] = disk.HeapSigTC
	data[1] = byte(len(w.columns)) //nolint:gosec
	binary.LittleEndian.PutUint16(data[2:4], end4)
	binary.LittleEndian.PutUint16(data[4:6], end2)
	binary.LittleEndian.PutUint16(data[6:8], end1)
	binary.LittleEndian.PutUint16(data[8:10], rowSize)
	binary.LittleEndian.PutUint32(data[10:14], uint32(rowBTHHID))
	binary.LittleEndian.PutUint32(data[14:18], uint32(rowMatrixHNID))
	copy(data[22:], w.buildColumnDescriptors())
	return data
}

func (w *TableWriter) RowCount() int {
	return len(w.rows)
}

func (w *TableWriter) ColumnCount() int {
	return len(w.columns)
}

func CreateHierarchyTable(format disk.PSTFormat) *TableWriter {
	w := NewTableWriter(format)
	w.AddColumn(PidTagLtpRowId, PropTypeInt32)
	w.AddColumn(PidTagDisplayName, PropTypeString)
	w.AddColumn(PidTagContentCount, PropTypeInt32)
	w.AddColumn(PidTagContentUnreadCount, PropTypeInt32)
	w.AddColumn(PidTagSubfolders, PropTypeBool)
	w.AddColumn(PidTagDepth, PropTypeInt32)
	return w
}

func CreateContentsTable(format disk.PSTFormat) *TableWriter {
	w := NewTableWriter(format)
	w.AddColumn(PidTagLtpRowId, PropTypeInt32)
	w.AddColumn(PidTagSubject, PropTypeString)
	w.AddColumn(PidTagMessageClass, PropTypeString)
	w.AddColumn(PidTagClientSubmitTime, PropTypeSysTime)
	w.AddColumn(PidTagMessageSize, PropTypeInt32)
	w.AddColumn(PidTagMessageFlags, PropTypeInt32)
	w.AddColumn(PidTagHasAttachments, PropTypeBool)
	return w
}

func CreateRecipientTable(format disk.PSTFormat) *TableWriter {
	w := NewTableWriter(format)
	w.AddColumn(PidTagLtpRowId, PropTypeInt32)
	w.AddColumn(PidTagRecipientType, PropTypeInt32)
	w.AddColumn(PidTagDisplayName, PropTypeString)
	w.AddColumn(PidTagEmailAddress, PropTypeString)
	w.AddColumn(PidTagAddressType, PropTypeString)
	return w
}

func CreateAttachmentTable(format disk.PSTFormat) *TableWriter {
	w := NewTableWriter(format)
	w.AddColumn(PidTagLtpRowId, PropTypeInt32)
	w.AddColumn(PidTagAttachNumber, PropTypeInt32)
	w.AddColumn(PidTagAttachMethod, PropTypeInt32)
	w.AddColumn(PidTagAttachFilename, PropTypeString)
	w.AddColumn(PidTagAttachSize, PropTypeInt32)
	w.AddColumn(PidTagAttachMimeTag, PropTypeString)
	return w
}
