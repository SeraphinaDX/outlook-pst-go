package disk

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"sort"
)

// AMapPage represents a single allocation map page.
// See [MS-PST] Section 2.2.2.7.2 - AMAPPAGE structure.
// Each AMap page tracks 253952 bytes (496 slots * 512 bytes/slot).
type AMapPage struct {
	Offset    uint64      // File offset of this AMap page
	BitMap    [496]byte   // Allocation bitmap (496 bytes = 3968 bits)
	Trailer   PageTrailer // Page trailer
	FreeSlots int         // Count of free 64-byte slots in this page's range
	format    PSTFormat
}

// AMapManager manages allocation map pages for a PST file.
// It tracks free space and allocates/frees blocks.
type AMapManager struct {
	pages     []*AMapPage
	format    PSTFormat
	fileSize  uint64
	freeBytes uint64 // Total free space in bytes

	// File reference for reading/writing
	reader io.ReaderAt

	// Pending allocations (not yet written to disk)
	pendingAllocs []allocation
}

// allocation represents a pending space allocation.
type allocation struct {
	offset uint64
	size   uint64
}

// NewAMapManager creates a new AMap manager from the PST header.
func NewAMapManager(r io.ReaderAt, h *Header) (*AMapManager, error) {
	m := &AMapManager{
		format:   h.Format,
		fileSize: h.Root.IBFileEOF,
		reader:   r,
	}

	// Load existing AMap pages
	if err := m.loadAMapPages(h); err != nil {
		return nil, fmt.Errorf("failed to load AMap pages: %w", err)
	}

	return m, nil
}

// loadAMapPages loads all AMap pages from the file.
func (m *AMapManager) loadAMapPages(h *Header) error {
	// First AMap page is always at FirstAMapPageLocation (0x4400)
	offset := uint64(FirstAMapPageLocation)

	// Calculate number of AMap pages needed for current file size
	// Each AMap page covers 253952 bytes (496 * 512)
	bytesPerAMap := uint64(496 * 512)

	for offset <= h.Root.IBAMapLast {
		page, err := m.readAMapPage(offset)
		if err != nil {
			return fmt.Errorf("failed to read AMap page at 0x%X: %w", offset, err)
		}
		m.pages = append(m.pages, page)

		// Calculate next AMap page offset
		// AMaps are spaced at bytesPerAMap intervals after accounting for header
		offset += bytesPerAMap
		if offset == FirstAMapPageLocation+bytesPerAMap {
			// First interval needs to account for initial file structures
			offset = FirstAMapPageLocation + bytesPerAMap
		}
	}

	// Calculate total free space
	m.calculateFreeSpace()

	return nil
}

// readAMapPage reads an AMap page from the file.
func (m *AMapManager) readAMapPage(offset uint64) (*AMapPage, error) {
	buf := make([]byte, PageSize)
	n, err := m.reader.ReadAt(buf, int64(offset)) //nolint:gosec // G115: PST file size bounded by format (max ~50GB)
	if err != nil && err != io.EOF {
		return nil, err
	}
	if n != PageSize {
		return nil, fmt.Errorf("incomplete AMap page read: got %d bytes", n)
	}

	page := &AMapPage{
		Offset: offset,
		format: m.format,
	}

	// Copy bitmap (first 496 bytes)
	copy(page.BitMap[:], buf[:496])

	// Parse trailer
	trailerOffset := PageSize - PageTrailerSizeUnicode
	if m.format == FormatANSI {
		trailerOffset = PageSize - PageTrailerSizeANSI
	}

	if m.format == FormatUnicode {
		page.Trailer = ParsePageTrailerUnicode(buf[trailerOffset:])
	} else {
		page.Trailer = ParsePageTrailerANSI(buf[trailerOffset:])
	}

	// Verify page type
	if page.Trailer.PageType != PageTypeAMap {
		return nil, fmt.Errorf("invalid AMap page type: got 0x%02X", page.Trailer.PageType)
	}

	// Count free slots
	page.countFreeSlots()

	return page, nil
}

// countFreeSlots counts the number of free 64-byte slots in this AMap page.
func (p *AMapPage) countFreeSlots() {
	p.FreeSlots = 0
	for _, b := range p.BitMap {
		// Count zero bits (free slots)
		for i := 0; i < 8; i++ {
			if b&(1<<i) == 0 {
				p.FreeSlots++
			}
		}
	}
}

// calculateFreeSpace calculates total free space across all AMap pages.
func (m *AMapManager) calculateFreeSpace() {
	m.freeBytes = 0
	for _, page := range m.pages {
		m.freeBytes += uint64(page.FreeSlots) * BytesPerSlot //nolint:gosec // G115: FreeSlots bounded by AMap page size
	}
}

// FreeSpace returns the total free space in bytes.
func (m *AMapManager) FreeSpace() uint64 {
	return m.freeBytes
}

// FileSize returns the current file size.
func (m *AMapManager) FileSize() uint64 {
	return m.fileSize
}

// Allocate allocates space for a block of the given size.
// Returns the file offset where the block should be written.
// The size should be the aligned block size including trailer.
func (m *AMapManager) Allocate(size uint64) (uint64, error) {
	if size == 0 {
		return 0, errors.New("cannot allocate zero bytes")
	}

	// Align size to 64-byte boundary
	alignedSize := AlignDisk(size)
	slotsNeeded := int(alignedSize / BytesPerSlot) //nolint:gosec // G115: slot count bounded by AMap page structure

	// Try to find contiguous free space in existing pages
	for _, page := range m.pages {
		offset, ok := page.findContiguousFree(slotsNeeded)
		if ok {
			// Mark slots as allocated
			page.markAllocated(offset, slotsNeeded)
			m.freeBytes -= alignedSize

			// Record pending allocation
			m.pendingAllocs = append(m.pendingAllocs, allocation{
				offset: offset,
				size:   alignedSize,
			})

			if end := offset + alignedSize; end > m.fileSize {
				m.fileSize = end
			}
			return offset, nil
		}
	}

	// No space found - need to extend file
	// For now, allocate at end of file
	newOffset := m.fileSize
	m.fileSize += alignedSize

	// Record pending allocation
	m.pendingAllocs = append(m.pendingAllocs, allocation{
		offset: newOffset,
		size:   alignedSize,
	})

	return newOffset, nil
}

// findContiguousFree finds contiguous free slots in the AMap page.
// Returns the file offset and whether found.
func (p *AMapPage) findContiguousFree(slotsNeeded int) (uint64, bool) {
	if slotsNeeded > p.FreeSlots {
		return 0, false
	}

	// Calculate the byte range this AMap covers
	baseOffset := p.Offset + PageSize // Data starts after the AMap page itself
	if p.Offset == FirstAMapPageLocation {
		// First AMap - data starts after initial file structures
		baseOffset = FirstAMapPageLocation + PageSize
	}

	// Search for contiguous free slots
	consecutiveFree := 0
	startSlot := -1

	for byteIdx := 0; byteIdx < len(p.BitMap); byteIdx++ {
		b := p.BitMap[byteIdx]
		for bitIdx := 0; bitIdx < 8; bitIdx++ {
			slotIdx := byteIdx*8 + bitIdx
			if b&(1<<bitIdx) == 0 {
				// Slot is free
				if consecutiveFree == 0 {
					startSlot = slotIdx
				}
				consecutiveFree++
				if consecutiveFree >= slotsNeeded {
					return baseOffset + uint64(startSlot)*BytesPerSlot, true
				}
			} else {
				// Slot is allocated - reset counter
				consecutiveFree = 0
				startSlot = -1
			}
		}
	}

	return 0, false
}

// markAllocated marks slots as allocated starting at the given offset.
func (p *AMapPage) markAllocated(offset uint64, slots int) {
	baseOffset := p.Offset + PageSize
	if p.Offset == FirstAMapPageLocation {
		baseOffset = FirstAMapPageLocation + PageSize
	}

	startSlot := int((offset - baseOffset) / BytesPerSlot) //nolint:gosec // G115: slot index bounded by AMap page bitmap

	for i := 0; i < slots; i++ {
		slotIdx := startSlot + i
		byteIdx := slotIdx / 8
		bitIdx := slotIdx % 8
		if byteIdx < len(p.BitMap) {
			p.BitMap[byteIdx] |= (1 << bitIdx)
			p.FreeSlots--
		}
	}
}

// Free frees previously allocated space.
func (m *AMapManager) Free(offset, size uint64) error {
	alignedSize := AlignDisk(size)
	slotsToFree := int(alignedSize / BytesPerSlot) //nolint:gosec // G115: slot count bounded by AMap page structure

	// Find the AMap page that covers this offset
	page := m.findPageForOffset(offset)
	if page == nil {
		return fmt.Errorf("no AMap page found for offset 0x%X", offset)
	}

	// Mark slots as free
	page.markFree(offset, slotsToFree)
	m.freeBytes += alignedSize

	return nil
}

// findPageForOffset finds the AMap page that covers the given offset.
func (m *AMapManager) findPageForOffset(offset uint64) *AMapPage {
	bytesPerAMap := uint64(496 * 512)

	for _, page := range m.pages {
		pageBase := page.Offset + PageSize
		if page.Offset == FirstAMapPageLocation {
			pageBase = FirstAMapPageLocation + PageSize
		}
		pageEnd := pageBase + bytesPerAMap

		if offset >= pageBase && offset < pageEnd {
			return page
		}
	}
	return nil
}

// markFree marks slots as free starting at the given offset.
func (p *AMapPage) markFree(offset uint64, slots int) {
	baseOffset := p.Offset + PageSize
	if p.Offset == FirstAMapPageLocation {
		baseOffset = FirstAMapPageLocation + PageSize
	}

	startSlot := int((offset - baseOffset) / BytesPerSlot) //nolint:gosec // G115: slot index bounded by AMap page bitmap

	for i := 0; i < slots; i++ {
		slotIdx := startSlot + i
		byteIdx := slotIdx / 8
		bitIdx := slotIdx % 8
		if byteIdx < len(p.BitMap) {
			p.BitMap[byteIdx] &^= (1 << bitIdx)
			p.FreeSlots++
		}
	}
}

// RebuildFromBTrees rebuilds the allocation map from the B-tree contents.
// This is used for recovery or validation.
func (m *AMapManager) RebuildFromBTrees(allocatedBlocks []BlockAllocation) error {
	// Reset all AMap pages to fully free
	for _, page := range m.pages {
		for i := range page.BitMap {
			page.BitMap[i] = 0
		}
		page.countFreeSlots()
	}

	// Sort allocations by offset for efficient processing
	sort.Slice(allocatedBlocks, func(i, j int) bool {
		return allocatedBlocks[i].Offset < allocatedBlocks[j].Offset
	})

	// Mark each block as allocated
	for _, block := range allocatedBlocks {
		page := m.findPageForOffset(block.Offset)
		if page == nil {
			// Block is beyond current AMap coverage - extend if needed
			continue
		}

		alignedSize := AlignDisk(uint64(block.Size))
		slots := int(alignedSize / BytesPerSlot) //nolint:gosec // G115: slot count bounded by block size limit
		page.markAllocated(block.Offset, slots)
	}

	// Recalculate free space
	m.calculateFreeSpace()

	return nil
}

// BlockAllocation represents an allocated block for AMap rebuilding.
type BlockAllocation struct {
	Offset uint64
	Size   uint16
}

// SerializeAMapPage serializes an AMap page for writing.
func (p *AMapPage) Serialize() ([]byte, error) {
	buf := make([]byte, PageSize)

	// Write bitmap
	copy(buf[:496], p.BitMap[:])

	// Compute and write trailer
	trailerOffset := PageSize - PageTrailerSizeUnicode
	trailerSize := PageTrailerSizeUnicode
	if p.format == FormatANSI {
		trailerOffset = PageSize - PageTrailerSizeANSI
		trailerSize = PageTrailerSizeANSI
	}

	// Compute CRC of page data (excluding trailer)
	crc := ComputeCRC(buf[:trailerOffset])

	// Compute signature
	sig := ComputeSignature(p.Trailer.BID, p.Offset)

	// Write trailer
	trailer := buf[trailerOffset:]
	trailer[0] = byte(PageTypeAMap)
	trailer[1] = byte(PageTypeAMap) // ptypeRepeat
	binary.LittleEndian.PutUint16(trailer[2:4], sig)

	if p.format == FormatUnicode {
		binary.LittleEndian.PutUint32(trailer[4:8], crc)
		binary.LittleEndian.PutUint64(trailer[8:16], p.Trailer.BID)
	} else {
		binary.LittleEndian.PutUint32(trailer[4:8], uint32(p.Trailer.BID)) //nolint:gosec // G115: ANSI format BID is 32-bit per MS-PST spec
		binary.LittleEndian.PutUint32(trailer[8:12], crc)
	}

	_ = trailerSize // Used for documentation

	return buf, nil
}

// WritePages serializes the current allocation maps to disk.
func (m *AMapManager) WritePages(w io.WriterAt) error {
	for _, page := range m.pages {
		data, err := page.Serialize()
		if err != nil {
			return fmt.Errorf("serialize AMap page at 0x%X: %w", page.Offset, err)
		}
		if _, err := w.WriteAt(data, int64(page.Offset)); err != nil { //nolint:gosec
			return fmt.Errorf("write AMap page at 0x%X: %w", page.Offset, err)
		}
	}
	return nil
}

// PendingAllocations returns the list of pending allocations.
func (m *AMapManager) PendingAllocations() []allocation {
	return m.pendingAllocs
}

// ClearPending clears the pending allocations list.
func (m *AMapManager) ClearPending() {
	m.pendingAllocs = nil
}

// CreateNewAMapManager creates an AMap manager for a new PST file.
func CreateNewAMapManager(format PSTFormat) *AMapManager {
	m := &AMapManager{
		format:   format,
		fileSize: uint64(FirstAMapPageLocation + PageSize), // Header + DList + AMap
	}

	// Create initial AMap page
	page := &AMapPage{
		Offset: FirstAMapPageLocation,
		format: format,
		Trailer: PageTrailer{
			PageType:       PageTypeAMap,
			PageTypeRepeat: PageTypeAMap,
			BID:            0, // AMap pages have BID 0
		},
	}

	// Mark initial file structures as allocated
	// Header: 0x0000 - HeaderSize
	// Reserved: HeaderSize - 0x4200
	// DList: 0x4200 - 0x4400
	// First AMap: 0x4400 - 0x4600

	// Calculate slots for initial structures
	initialSlots := int((FirstAMapPageLocation + PageSize) / BytesPerSlot)
	for i := 0; i < initialSlots && i < 496*8; i++ {
		byteIdx := i / 8
		bitIdx := i % 8
		page.BitMap[byteIdx] |= (1 << bitIdx)
	}

	page.countFreeSlots()
	m.pages = append(m.pages, page)
	m.calculateFreeSpace()

	return m
}
