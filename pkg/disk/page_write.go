package disk

import (
	"encoding/binary"
	"fmt"
)

// BuildPage creates a complete page with data and trailer.
// The returned bytes are exactly PageSize (512 bytes).
// Parameters:
//   - data: the page payload (will be padded to fill page)
//   - pageType: type of page (NBT, BBT, AMap, etc.)
//   - bid: block/page ID
//   - offset: file offset (for signature computation)
//   - format: PST format
func BuildPage(data []byte, pageType PageType, bid uint64, offset uint64, format PSTFormat) ([]byte, error) {
	trailerSize := PageTrailerSizeUnicode
	if format == FormatANSI {
		trailerSize = PageTrailerSizeANSI
	}

	maxDataSize := PageSize - trailerSize
	if len(data) > maxDataSize {
		return nil, fmt.Errorf("page data too large: %d bytes (max %d)", len(data), maxDataSize)
	}

	// Allocate full page
	buf := make([]byte, PageSize)

	// Copy data (rest is zero-padded)
	copy(buf, data)

	// Compute CRC of page data (excluding trailer)
	crc := ComputeCRC(buf[:PageSize-trailerSize])

	// Compute signature
	sig := ComputeSignature(bid, offset)

	// Write trailer
	trailer := buf[PageSize-trailerSize:]
	trailer[0] = byte(pageType)
	trailer[1] = byte(pageType) // ptypeRepeat

	binary.LittleEndian.PutUint16(trailer[2:4], sig)

	if format == FormatUnicode {
		// Unicode: wSig(2) + dwCRC(4) + bid(8)
		binary.LittleEndian.PutUint32(trailer[4:8], crc)
		binary.LittleEndian.PutUint64(trailer[8:16], bid)
	} else {
		// ANSI: wSig(2) + bid(4) + dwCRC(4)
		binary.LittleEndian.PutUint32(trailer[4:8], uint32(bid)) //nolint:gosec // G115: ANSI BID is 32-bit
		binary.LittleEndian.PutUint32(trailer[8:12], crc)
	}

	return buf, nil
}

// BuildBTPage creates a B-tree page (NBT or BBT).
func BuildBTPage(page *BTPage, format PSTFormat) ([]byte, error) {
	trailerSize := PageTrailerSizeUnicode
	if format == FormatANSI {
		trailerSize = PageTrailerSizeANSI
	}

	// Calculate maximum entry area. Unicode BTPAGE has a 4-byte
	// dwPadding field between the 4-byte BTPAGE header and the trailer.
	paddingSize := 0
	if format == FormatUnicode {
		paddingSize = 4
	}
	dataAreaSize := PageSize - trailerSize - 4 - paddingSize

	// Serialize entries first
	var entriesData []byte
	var entrySize int

	if page.IsLeaf() {
		if page.Trailer.PageType == PageTypeNBT {
			entriesData, entrySize = serializeNBTLeafEntries(page.NBTEntries, format)
		} else {
			entriesData, entrySize = serializeBBTLeafEntries(page.BBTEntries, format)
		}
	} else {
		entriesData, entrySize = serializeNonleafEntries(page.NonleafEntries, format)
	}

	if len(entriesData) > dataAreaSize {
		return nil, fmt.Errorf("too many entries for page: %d bytes (max %d)", len(entriesData), dataAreaSize)
	}

	// Build page data
	data := make([]byte, PageSize-trailerSize)

	// Copy entries at start
	copy(data, entriesData)

	// Write the BTPAGE header immediately after the entry area. Unicode
	// reserves an additional 4-byte dwPadding field before PAGETRAILER.
	headerOffset := PageSize - trailerSize - 4 - paddingSize
	numEntries := len(page.NBTEntries)
	if page.Trailer.PageType == PageTypeBBT {
		numEntries = len(page.BBTEntries)
	}
	if !page.IsLeaf() {
		numEntries = len(page.NonleafEntries)
	}

	// Calculate max entries based on entry size
	maxEntries := dataAreaSize / entrySize
	if maxEntries > 255 {
		maxEntries = 255
	}

	data[headerOffset] = byte(numEntries)   //nolint:gosec // G115: bounded by max 255
	data[headerOffset+1] = byte(maxEntries) //nolint:gosec // G115: bounded by max 255
	data[headerOffset+2] = byte(entrySize)  //nolint:gosec // G115: entry size is small constant
	data[headerOffset+3] = byte(page.Level) //nolint:gosec // G115: B-tree level bounded by depth limit

	// Build complete page with trailer
	return BuildPage(data, page.Trailer.PageType, page.Trailer.BID, page.Trailer.BID, format)
}

// serializeNBTLeafEntries serializes NBT leaf entries.
func serializeNBTLeafEntries(entries []NBTLeafEntry, format PSTFormat) ([]byte, int) {
	entrySize := NBTLeafEntrySizeUnicode
	if format == FormatANSI {
		entrySize = NBTLeafEntrySizeANSI
	}

	data := make([]byte, len(entries)*entrySize)

	for i, entry := range entries {
		offset := i * entrySize
		if format == FormatUnicode {
			// Unicode entry (32 bytes):
			// nid: 0-8 (8 bytes)
			// bidData: 8-16 (8 bytes)
			// bidSub: 16-24 (8 bytes)
			// nidParent: 24-28 (4 bytes)
			// dwPadding: 28-32 (4 bytes)
			binary.LittleEndian.PutUint64(data[offset:offset+8], entry.NID)
			binary.LittleEndian.PutUint64(data[offset+8:offset+16], entry.DataBID)
			binary.LittleEndian.PutUint64(data[offset+16:offset+24], entry.SubBID)
			binary.LittleEndian.PutUint32(data[offset+24:offset+28], entry.ParentNID)
		} else {
			// ANSI entry (16 bytes):
			// nid: 0-4 (4 bytes)
			// bidData: 4-8 (4 bytes)
			// bidSub: 8-12 (4 bytes)
			// nidParent: 12-16 (4 bytes)
			binary.LittleEndian.PutUint32(data[offset:offset+4], uint32(entry.NID))       //nolint:gosec // G115: ANSI NID is 32-bit
			binary.LittleEndian.PutUint32(data[offset+4:offset+8], uint32(entry.DataBID)) //nolint:gosec // G115: ANSI BID is 32-bit
			binary.LittleEndian.PutUint32(data[offset+8:offset+12], uint32(entry.SubBID)) //nolint:gosec // G115: ANSI BID is 32-bit
			binary.LittleEndian.PutUint32(data[offset+12:offset+16], entry.ParentNID)
		}
	}

	return data, entrySize
}

// serializeBBTLeafEntries serializes BBT leaf entries.
func serializeBBTLeafEntries(entries []BBTLeafEntry, format PSTFormat) ([]byte, int) {
	entrySize := BBTLeafEntrySizeUnicode
	if format == FormatANSI {
		entrySize = BBTLeafEntrySizeANSI
	}

	data := make([]byte, len(entries)*entrySize)

	for i, entry := range entries {
		offset := i * entrySize
		if format == FormatUnicode {
			// Unicode entry (24 bytes):
			// bref.bid: 0-8 (8 bytes)
			// bref.ib: 8-16 (8 bytes)
			// cb: 16-18 (2 bytes)
			// cRef: 18-20 (2 bytes)
			// dwPadding: 20-24 (4 bytes)
			binary.LittleEndian.PutUint64(data[offset:offset+8], entry.BRef.BID)
			binary.LittleEndian.PutUint64(data[offset+8:offset+16], entry.BRef.IB)
			binary.LittleEndian.PutUint16(data[offset+16:offset+18], entry.Size)
			binary.LittleEndian.PutUint16(data[offset+18:offset+20], entry.RefCount)
		} else {
			// ANSI entry (12 bytes):
			// bref.bid: 0-4 (4 bytes)
			// bref.ib: 4-8 (4 bytes)
			// cb: 8-10 (2 bytes)
			// cRef: 10-12 (2 bytes)
			binary.LittleEndian.PutUint32(data[offset:offset+4], uint32(entry.BRef.BID))  //nolint:gosec // G115: ANSI BID is 32-bit
			binary.LittleEndian.PutUint32(data[offset+4:offset+8], uint32(entry.BRef.IB)) //nolint:gosec // G115: ANSI offset is 32-bit
			binary.LittleEndian.PutUint16(data[offset+8:offset+10], entry.Size)
			binary.LittleEndian.PutUint16(data[offset+10:offset+12], entry.RefCount)
		}
	}

	return data, entrySize
}

// serializeNonleafEntries serializes B-tree non-leaf entries.
func serializeNonleafEntries(entries []BTNonleafEntry, format PSTFormat) ([]byte, int) {
	entrySize := 24 // Unicode: key(8) + bref(16)
	if format == FormatANSI {
		entrySize = 12 // ANSI: key(4) + bref(8)
	}

	data := make([]byte, len(entries)*entrySize)

	for i, entry := range entries {
		offset := i * entrySize
		if format == FormatUnicode {
			// Unicode non-leaf entry (24 bytes):
			// key: 0-8 (8 bytes)
			// bref.bid: 8-16 (8 bytes)
			// bref.ib: 16-24 (8 bytes)
			binary.LittleEndian.PutUint64(data[offset:offset+8], entry.Key)
			binary.LittleEndian.PutUint64(data[offset+8:offset+16], entry.Ref.BID)
			binary.LittleEndian.PutUint64(data[offset+16:offset+24], entry.Ref.IB)
		} else {
			// ANSI non-leaf entry (12 bytes):
			// key: 0-4 (4 bytes)
			// bref.bid: 4-8 (4 bytes)
			// bref.ib: 8-12 (4 bytes)
			binary.LittleEndian.PutUint32(data[offset:offset+4], uint32(entry.Key))       //nolint:gosec // G115: ANSI key is 32-bit
			binary.LittleEndian.PutUint32(data[offset+4:offset+8], uint32(entry.Ref.BID)) //nolint:gosec // G115: ANSI BID is 32-bit
			binary.LittleEndian.PutUint32(data[offset+8:offset+12], uint32(entry.Ref.IB)) //nolint:gosec // G115: ANSI offset is 32-bit
		}
	}

	return data, entrySize
}

// NewNBTLeafPage creates a new NBT leaf page.
func NewNBTLeafPage(entries []NBTLeafEntry, bid uint64, format PSTFormat) *BTPage {
	return &BTPage{
		Trailer: PageTrailer{
			PageType:       PageTypeNBT,
			PageTypeRepeat: PageTypeNBT,
			BID:            bid,
		},
		Header: BTPageHeader{
			Level: 0,
		},
		Level:      0,
		NBTEntries: entries,
	}
}

// NewBBTLeafPage creates a new BBT leaf page.
func NewBBTLeafPage(entries []BBTLeafEntry, bid uint64, format PSTFormat) *BTPage {
	return &BTPage{
		Trailer: PageTrailer{
			PageType:       PageTypeBBT,
			PageTypeRepeat: PageTypeBBT,
			BID:            bid,
		},
		Header: BTPageHeader{
			Level: 0,
		},
		Level:      0,
		BBTEntries: entries,
	}
}

// NewBTNonleafPage creates a new B-tree non-leaf page.
func NewBTNonleafPage(entries []BTNonleafEntry, level int, pageType PageType, bid uint64) *BTPage {
	return &BTPage{
		Trailer: PageTrailer{
			PageType:       pageType,
			PageTypeRepeat: pageType,
			BID:            bid,
		},
		Header: BTPageHeader{
			Level: byte(level), //nolint:gosec // G115: B-tree level bounded by depth limit
		},
		Level:          level,
		NonleafEntries: entries,
	}
}

// BuildDListPage creates a Density List page.
// DList pages track free space density across AMap pages for efficient allocation.
func BuildDListPage(entries []DListEntry, bid uint64, offset uint64, format PSTFormat) ([]byte, error) {
	trailerSize := PageTrailerSizeUnicode
	if format == FormatANSI {
		trailerSize = PageTrailerSizeANSI
	}

	// DList page structure:
	// bFlags: 1 byte (typically 0)
	// cEntDList: 1 byte (number of entries)
	// dwPadding: 2 bytes (Unicode) or 0 bytes (ANSI)
	// rgdlistent: array of entries (4 bytes each)

	headerSize := 4
	if format == FormatANSI {
		headerSize = 2
	}

	maxEntries := (PageSize - trailerSize - headerSize) / 4
	if len(entries) > maxEntries {
		return nil, fmt.Errorf("too many DList entries: %d (max %d)", len(entries), maxEntries)
	}

	data := make([]byte, PageSize-trailerSize)
	data[0] = 0                  // bFlags
	data[1] = byte(len(entries)) //nolint:gosec // G115: entry count bounded by DList capacity

	// Write entries
	pos := headerSize
	for _, entry := range entries {
		// Each entry is 4 bytes:
		// dwPageNum: 20 bits
		// dwFreeSlots: 12 bits
		val := (entry.PageNum & 0xFFFFF) | ((entry.FreeSlots & 0xFFF) << 20)
		binary.LittleEndian.PutUint32(data[pos:pos+4], val)
		pos += 4
	}

	return BuildPage(data, PageTypeDList, bid, offset, format)
}

// DListEntry represents an entry in the Density List.
type DListEntry struct {
	PageNum   uint32 // AMap page number (20 bits)
	FreeSlots uint32 // Free slots in that page (12 bits)
}

// ValidatePageTrailer validates a page trailer against computed values.
func ValidatePageTrailer(pageData []byte, trailer *PageTrailer, offset uint64, format PSTFormat) error {
	trailerSize := PageTrailerSizeUnicode
	if format == FormatANSI {
		trailerSize = PageTrailerSizeANSI
	}

	// Verify page types match
	if trailer.PageType != trailer.PageTypeRepeat {
		return fmt.Errorf("page type mismatch: %v != %v", trailer.PageType, trailer.PageTypeRepeat)
	}

	// Compute expected CRC
	expectedCRC := ComputeCRC(pageData[:PageSize-trailerSize])
	if trailer.CRC != expectedCRC {
		return fmt.Errorf("CRC mismatch: got 0x%08X, expected 0x%08X", trailer.CRC, expectedCRC)
	}

	// Compute expected signature
	expectedSig := ComputeSignature(trailer.BID, offset)
	if trailer.Signature != expectedSig {
		return fmt.Errorf("signature mismatch: got 0x%04X, expected 0x%04X", trailer.Signature, expectedSig)
	}

	return nil
}
