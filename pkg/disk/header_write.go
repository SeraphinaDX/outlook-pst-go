package disk

import (
	"encoding/binary"
	"fmt"
	"io"
)

// SerializeHeader serializes the header to bytes for writing.
// It automatically selects Unicode or ANSI format based on h.Format.
func SerializeHeader(h *Header) ([]byte, error) {
	if h.Format == FormatUnicode {
		return SerializeHeaderUnicode(h)
	}
	return SerializeHeaderANSI(h)
}

// SerializeHeaderUnicode serializes a Unicode format header.
func SerializeHeaderUnicode(h *Header) ([]byte, error) {
	buf := make([]byte, HeaderSizeUnicode)

	binary.LittleEndian.PutUint32(buf[0:4], h.DWMagic)
	binary.LittleEndian.PutUint32(buf[4:8], h.DWCRCPartial)
	binary.LittleEndian.PutUint16(buf[8:10], h.WMagicClient)
	binary.LittleEndian.PutUint16(buf[10:12], h.WVer)
	binary.LittleEndian.PutUint16(buf[12:14], h.WVerClient)
	buf[14] = h.BPlatformCreate
	buf[15] = h.BPlatformAccess
	binary.LittleEndian.PutUint32(buf[16:20], h.DWOpenDBID)
	binary.LittleEndian.PutUint32(buf[20:24], h.DWOpenClaimID)

	// bidUnused: 0x18..0x20
	binary.LittleEndian.PutUint64(buf[32:40], h.BidNextP)
	binary.LittleEndian.PutUint32(buf[40:44], h.DWUnique)
	// rgnid[32]: 0x2C..0xAC
	// qwUnused: 0xAC..0xB4

	// ROOT is 72 bytes at 0xB4.
	serializeRootUnicode(buf[180:252], &h.Root)
	// dwAlign: 0xFC..0x100
	// rgbFM: 0x100..0x180
	// rgbFP: 0x180..0x200

	buf[512] = 0x80
	buf[513] = byte(h.BCryptMethod)
	// rgbReserved: 0x202..0x204
	binary.LittleEndian.PutUint64(buf[516:524], h.BidNextB)
	// dwCRCFull at 0x20C is filled after CRC calculation.

	// [MS-PST] defines dwCRCPartial over 471 bytes starting at 0x08.
	h.DWCRCPartial = ComputeCRC(buf[8:479])
	binary.LittleEndian.PutUint32(buf[4:8], h.DWCRCPartial)

	// dwCRCFull covers 516 bytes starting at 0x08, ending immediately
	// before dwCRCFull itself at 0x20C.
	h.DWCRCFull = ComputeCRC(buf[8:524])
	binary.LittleEndian.PutUint32(buf[524:528], h.DWCRCFull)

	return buf, nil
}

// SerializeHeaderANSI serializes an ANSI format header (512 bytes).
func SerializeHeaderANSI(h *Header) ([]byte, error) {
	buf := make([]byte, HeaderSizeANSI)

	// Common header fields (0-24)
	binary.LittleEndian.PutUint32(buf[0:4], h.DWMagic)
	binary.LittleEndian.PutUint32(buf[4:8], h.DWCRCPartial)
	binary.LittleEndian.PutUint16(buf[8:10], h.WMagicClient)
	binary.LittleEndian.PutUint16(buf[10:12], h.WVer)
	binary.LittleEndian.PutUint16(buf[12:14], h.WVerClient)
	buf[14] = h.BPlatformCreate
	buf[15] = h.BPlatformAccess
	binary.LittleEndian.PutUint32(buf[16:20], h.DWOpenDBID)
	binary.LittleEndian.PutUint32(buf[20:24], h.DWOpenClaimID)

	// bidNextB: 24-28 (4 bytes) - in ANSI, this comes earlier
	binary.LittleEndian.PutUint32(buf[24:28], uint32(h.BidNextB)) //nolint:gosec // G115: ANSI BID is 32-bit
	// bidNextP: 28-32 (4 bytes)
	binary.LittleEndian.PutUint32(buf[28:32], uint32(h.BidNextP)) //nolint:gosec // G115: ANSI BID is 32-bit
	// dwUnique: 32-36 (4 bytes)
	binary.LittleEndian.PutUint32(buf[32:36], h.DWUnique)

	// rgnid[32]: 36-164 (128 bytes) - node ID counters, write zeros

	// Root structure: 164-204 (40 bytes)
	serializeRootANSI(buf[164:], &h.Root)

	// rgbFM: 204-332 (128 bytes) - deprecated
	// rgbFP: 332-460 (128 bytes) - deprecated

	// bSentinel: 460 (1 byte)
	buf[460] = 0x80

	// bCryptMethod: 461 (1 byte)
	buf[461] = byte(h.BCryptMethod)

	// rgbReserved: 462-464 (2 bytes)
	// ullReserved: 464-472 (8 bytes)
	// Additional reserved fields

	// Compute CRC (bytes 8-471)
	h.DWCRCPartial = ComputeCRC(buf[8:472])
	binary.LittleEndian.PutUint32(buf[4:8], h.DWCRCPartial)

	return buf, nil
}

// serializeRootUnicode serializes the 72-byte Unicode ROOT structure.
func serializeRootUnicode(buf []byte, r *Root) {
	binary.LittleEndian.PutUint32(buf[0:4], r.COrphans)
	binary.LittleEndian.PutUint64(buf[4:12], r.IBFileEOF)
	binary.LittleEndian.PutUint64(buf[12:20], r.IBAMapLast)
	binary.LittleEndian.PutUint64(buf[20:28], r.CBAMapFree)
	binary.LittleEndian.PutUint64(buf[28:36], r.CBPMapFree)
	binary.LittleEndian.PutUint64(buf[36:44], r.BRefNBT.BID)
	binary.LittleEndian.PutUint64(buf[44:52], r.BRefNBT.IB)
	binary.LittleEndian.PutUint64(buf[52:60], r.BRefBBT.BID)
	binary.LittleEndian.PutUint64(buf[60:68], r.BRefBBT.IB)
	buf[68] = r.FAMapValid
	buf[69] = r.BARVec
	binary.LittleEndian.PutUint16(buf[70:72], r.CARVec)
}

// serializeRootANSI serializes the Root structure for ANSI format.
func serializeRootANSI(buf []byte, r *Root) {
	// Root structure for ANSI (40 bytes):
	// cOrphans: 0-4 (4 bytes)
	binary.LittleEndian.PutUint32(buf[0:4], r.COrphans)
	// ibFileEOF: 4-8 (4 bytes)
	binary.LittleEndian.PutUint32(buf[4:8], uint32(r.IBFileEOF)) //nolint:gosec // G115: ANSI file size is 32-bit
	// ibAMapLast: 8-12 (4 bytes)
	binary.LittleEndian.PutUint32(buf[8:12], uint32(r.IBAMapLast)) //nolint:gosec // G115: ANSI offset is 32-bit
	// cbAMapFree: 12-16 (4 bytes)
	binary.LittleEndian.PutUint32(buf[12:16], uint32(r.CBAMapFree)) //nolint:gosec // G115: ANSI size is 32-bit
	// cbPMapFree: 16-20 (4 bytes)
	binary.LittleEndian.PutUint32(buf[16:20], uint32(r.CBPMapFree)) //nolint:gosec // G115: ANSI size is 32-bit
	// brefNBT: 20-28 (8 bytes) - bid(4) + ib(4)
	binary.LittleEndian.PutUint32(buf[20:24], uint32(r.BRefNBT.BID)) //nolint:gosec // G115: ANSI BID is 32-bit
	binary.LittleEndian.PutUint32(buf[24:28], uint32(r.BRefNBT.IB))  //nolint:gosec // G115: ANSI offset is 32-bit
	// brefBBT: 28-36 (8 bytes)
	binary.LittleEndian.PutUint32(buf[28:32], uint32(r.BRefBBT.BID)) //nolint:gosec // G115: ANSI BID is 32-bit
	binary.LittleEndian.PutUint32(buf[32:36], uint32(r.BRefBBT.IB))  //nolint:gosec // G115: ANSI offset is 32-bit
	// fAMapValid: 36 (1 byte)
	buf[36] = r.FAMapValid
	// bARVec: 37 (1 byte)
	buf[37] = r.BARVec
	// cARVec: 38-40 (2 bytes)
	binary.LittleEndian.PutUint16(buf[38:40], r.CARVec)
}

// WriteHeader writes the header to the given writer at position 0.
func WriteHeader(w io.WriterAt, h *Header) error {
	data, err := SerializeHeader(h)
	if err != nil {
		return fmt.Errorf("failed to serialize header: %w", err)
	}

	n, err := w.WriteAt(data, 0)
	if err != nil {
		return fmt.Errorf("failed to write header: %w", err)
	}
	if n != len(data) {
		return fmt.Errorf("incomplete header write: wrote %d bytes, expected %d", n, len(data))
	}

	return nil
}

// NewHeader creates a new PST header with default values.
// This is used when creating a new PST file.
func NewHeader(format PSTFormat, clientType uint16) *Header {
	h := &Header{
		DWMagic:         PSTMagic,
		WMagicClient:    clientType,
		BPlatformCreate: 0x01, // Win32
		BPlatformAccess: 0x01, // Win32
		Format:          format,
		BCryptMethod:    CryptMethodPermute, // Default encryption
	}

	if format == FormatUnicode {
		h.WVer = DatabaseFormatUnicodeMax // Version 23
		h.WVerClient = 19                 // Standard client version
	} else {
		h.WVer = DatabaseFormatANSIMax // Version 15
		h.WVerClient = 19
	}

	// Initialize root with valid AMap status
	h.Root.FAMapValid = byte(AMapStatusValid)

	return h
}

// SetAMapStatus updates the AMap validity status in the header.
// This is used during the two-phase commit protocol.
func (h *Header) SetAMapStatus(status AMapStatus) {
	h.Root.FAMapValid = byte(status)
}

// GetAMapStatus returns the current AMap validity status.
func (h *Header) GetAMapStatus() AMapStatus {
	return AMapStatus(h.Root.FAMapValid)
}

// UpdateBTreeRoots updates the B-tree root references in the header.
func (h *Header) UpdateBTreeRoots(nbtRef, bbtRef BlockReference) {
	h.Root.BRefNBT = nbtRef
	h.Root.BRefBBT = bbtRef
}

// UpdateFileSize updates the file size in the header.
func (h *Header) UpdateFileSize(size uint64) {
	h.Root.IBFileEOF = size
}

// UpdateNextBlockID updates the next block ID counter.
func (h *Header) UpdateNextBlockID(bid uint64) {
	h.BidNextB = bid
}

// IncrementUnique increments the unique counter.
// This should be called on each modification.
func (h *Header) IncrementUnique() {
	h.DWUnique++
}
