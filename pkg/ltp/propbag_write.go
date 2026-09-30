package ltp

import (
	"encoding/binary"
	"fmt"
	"math"
	"time"
	"unicode/utf16"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// PropertyBagWriter builds a Property Context structure for writing.
// See [MS-PST] Section 2.3.3 for the Property Context specification.
type PropertyBagWriter struct {
	heap       *HeapWriter
	bth        *BTHWriter
	format     disk.PSTFormat
	properties map[PropID]*propertyData
}

// propertyData holds property information during building.
type propertyData struct {
	propType PropType
	value    []byte          // For fixed-size: the value; for variable: the data
	hnid     util.HeapNodeID // HID or subnode NID for indirect data
}

// NewPropertyBagWriter creates a new property bag writer.
func NewPropertyBagWriter(format disk.PSTFormat) *PropertyBagWriter {
	heap := CreatePropertyContextHeap(format)
	bth := CreatePropertyContextBTH(heap, format)

	return &PropertyBagWriter{
		heap:       heap,
		bth:        bth,
		format:     format,
		properties: make(map[PropID]*propertyData),
	}
}

// SetInt16 sets an int16 property.
func (w *PropertyBagWriter) SetInt16(id PropID, value int16) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data, uint16(value)) //nolint:gosec // G115: signed/unsigned reinterpretation, same bit width
	return w.setFixedProperty(id, PropTypeInt16, data)
}

// SetInt32 sets an int32 property.
func (w *PropertyBagWriter) SetInt32(id PropID, value int32) error {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint32(data, uint32(value)) //nolint:gosec // G115: signed/unsigned reinterpretation, same bit width
	return w.setFixedProperty(id, PropTypeInt32, data)
}

// SetInt64 sets an int64 property.
func (w *PropertyBagWriter) SetInt64(id PropID, value int64) error {
	// Int64 is variable-size in PC (stored as HID reference)
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data, uint64(value)) //nolint:gosec // G115: signed/unsigned reinterpretation, same bit width
	return w.setVariableProperty(id, PropTypeInt64, data)
}

// SetBool sets a boolean property.
func (w *PropertyBagWriter) SetBool(id PropID, value bool) error {
	data := make([]byte, 2)
	if value {
		binary.LittleEndian.PutUint16(data, 1)
	}
	return w.setFixedProperty(id, PropTypeBool, data)
}

// SetString sets a Unicode string property.
func (w *PropertyBagWriter) SetString(id PropID, value string) error {
	// Convert to UTF-16LE with null terminator
	data := encodeUTF16LE(value)
	return w.setVariableProperty(id, PropTypeString, data)
}

// SetString8 sets an ANSI string property.
func (w *PropertyBagWriter) SetString8(id PropID, value string) error {
	// For ANSI, just use the bytes with null terminator
	data := append([]byte(value), 0)
	return w.setVariableProperty(id, PropTypeString8, data)
}

// SetBinary sets a binary property.
func (w *PropertyBagWriter) SetBinary(id PropID, value []byte) error {
	return w.setVariableProperty(id, PropTypeBinary, value)
}

// SetTime sets a FILETIME property.
func (w *PropertyBagWriter) SetTime(id PropID, value time.Time) error {
	ft := TimeToFileTime(value)
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data, ft)
	return w.setVariableProperty(id, PropTypeSysTime, data)
}

// SetFloat32 sets a float32 property.
func (w *PropertyBagWriter) SetFloat32(id PropID, value float32) error {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint32(data, uint32FromFloat32(value))
	return w.setFixedProperty(id, PropTypeFloat32, data)
}

// SetFloat64 sets a float64 property.
func (w *PropertyBagWriter) SetFloat64(id PropID, value float64) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data, uint64FromFloat64(value))
	return w.setVariableProperty(id, PropTypeFloat64, data)
}

// SetGUID sets a GUID property.
func (w *PropertyBagWriter) SetGUID(id PropID, value util.GUID) error {
	return w.setVariableProperty(id, PropTypeGUID, value.Bytes())
}

// Delete removes a property.
func (w *PropertyBagWriter) Delete(id PropID) {
	delete(w.properties, id)
}

// setFixedProperty sets a fixed-size property.
func (w *PropertyBagWriter) setFixedProperty(id PropID, propType PropType, value []byte) error {
	w.properties[id] = &propertyData{
		propType: propType,
		value:    value,
	}
	return nil
}

// setVariableProperty sets a variable-size property.
func (w *PropertyBagWriter) setVariableProperty(id PropID, propType PropType, value []byte) error {
	w.properties[id] = &propertyData{
		propType: propType,
		value:    value,
	}
	return nil
}

// Build builds the property context and returns the heap data.
// Properties larger than a single Heap-on-Node allocation require BuildWithSubnodes.
func (w *PropertyBagWriter) Build() ([]byte, error) {
	return w.build(nil, nil)
}

// BuildWithSubnodes builds the property context and stores oversized property
// values as LTP subnodes. nextIndex is advanced for every generated subnode NID.
func (w *PropertyBagWriter) BuildWithSubnodes(subnodes *ndb.SubnodeBuilder, nextIndex *uint32) ([]byte, error) {
	if subnodes == nil || nextIndex == nil {
		return nil, fmt.Errorf("subnode builder and index are required")
	}
	return w.build(subnodes, nextIndex)
}

func (w *PropertyBagWriter) build(subnodes *ndb.SubnodeBuilder, nextIndex *uint32) ([]byte, error) {
	for id, prop := range w.properties {
		if prop.propType.IsFixedSize() && len(prop.value) <= 4 {
			continue
		}

		if len(prop.value) > disk.HeapMaxAllocSize {
			if subnodes == nil {
				return nil, fmt.Errorf("failed to allocate property 0x%04X: allocation too large: %d bytes (max %d)", id, len(prop.value), disk.HeapMaxAllocSize)
			}
			nid := util.MakeNID(util.NIDTypeLTP, *nextIndex)
			*nextIndex++
			if err := subnodes.AddSubnode(nid, prop.value); err != nil {
				return nil, fmt.Errorf("failed to store property 0x%04X as subnode: %w", id, err)
			}
			prop.hnid = util.HeapNodeID(nid)
			continue
		}

		hid, err := w.heap.Allocate(prop.value)
		if err != nil {
			return nil, fmt.Errorf("failed to allocate property 0x%04X: %w", id, err)
		}
		prop.hnid = util.HeapNodeID(hid)
	}

	// Second pass: build BTH entries
	for id, prop := range w.properties {
		// Property entry: type(2) + HNID(4)
		entry := make([]byte, 6)
		binary.LittleEndian.PutUint16(entry[0:2], uint16(prop.propType))

		if prop.propType.IsFixedSize() && len(prop.value) <= 4 {
			// Fixed-size: store value inline (padded to 4 bytes)
			var hnid uint32
			switch len(prop.value) {
			case 1:
				hnid = uint32(prop.value[0])
			case 2:
				hnid = uint32(binary.LittleEndian.Uint16(prop.value))
			case 4:
				hnid = binary.LittleEndian.Uint32(prop.value)
			}
			binary.LittleEndian.PutUint32(entry[2:6], hnid)
		} else {
			// Indirect value: store either a HID or a subnode NID.
			binary.LittleEndian.PutUint32(entry[2:6], uint32(prop.hnid))
		}

		if err := w.bth.InsertUint16Key(uint16(id), entry); err != nil {
			return nil, fmt.Errorf("failed to insert property 0x%04X: %w", id, err)
		}
	}

	// Build BTH
	_, err := w.bth.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build BTH: %w", err)
	}

	// Build heap
	return w.heap.Build()
}

// PropertyCount returns the number of properties.
func (w *PropertyBagWriter) PropertyCount() int {
	return len(w.properties)
}

// encodeUTF16LE encodes a string to UTF-16LE with null terminator.
func encodeUTF16LE(s string) []byte {
	runes := []rune(s)
	u16s := utf16.Encode(runes)
	// Add null terminator
	u16s = append(u16s, 0)

	// Convert to bytes
	buf := make([]byte, len(u16s)*2)
	for i, u := range u16s {
		binary.LittleEndian.PutUint16(buf[i*2:], u)
	}
	return buf
}

// Helper functions for float conversion.
func uint32FromFloat32(f float32) uint32 {
	return math.Float32bits(f)
}

func uint64FromFloat64(f float64) uint64 {
	return math.Float64bits(f)
}

// CreateMessagePropertyBag creates a property bag with common message properties.
func CreateMessagePropertyBag(format disk.PSTFormat, subject, body string, sentTime time.Time) (*PropertyBagWriter, error) {
	w := NewPropertyBagWriter(format)

	// Set message class
	if err := w.SetString(PidTagMessageClass, "IPM.Note"); err != nil {
		return nil, err
	}

	// Set subject
	if err := w.SetString(PidTagSubject, subject); err != nil {
		return nil, err
	}

	// Set body
	if err := w.SetString(PidTagBody, body); err != nil {
		return nil, err
	}

	// Set sent time
	if err := w.SetTime(PidTagClientSubmitTime, sentTime); err != nil {
		return nil, err
	}

	// Set creation time
	now := time.Now()
	if err := w.SetTime(PidTagCreationTime, now); err != nil {
		return nil, err
	}
	if err := w.SetTime(PidTagLastModificationTime, now); err != nil {
		return nil, err
	}

	return w, nil
}

// CreateFolderPropertyBag creates a property bag for a folder.
func CreateFolderPropertyBag(format disk.PSTFormat, displayName string) (*PropertyBagWriter, error) {
	w := NewPropertyBagWriter(format)

	// Set display name
	if err := w.SetString(PidTagDisplayName, displayName); err != nil {
		return nil, err
	}

	// Set content count (initially 0)
	if err := w.SetInt32(PidTagContentCount, 0); err != nil {
		return nil, err
	}

	// Set unread count (initially 0)
	if err := w.SetInt32(PidTagContentUnreadCount, 0); err != nil {
		return nil, err
	}

	// Set subfolders flag
	if err := w.SetBool(PidTagSubfolders, false); err != nil {
		return nil, err
	}

	// Set timestamps
	now := time.Now()
	if err := w.SetTime(PidTagCreationTime, now); err != nil {
		return nil, err
	}
	if err := w.SetTime(PidTagLastModificationTime, now); err != nil {
		return nil, err
	}

	return w, nil
}
