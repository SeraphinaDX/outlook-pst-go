package outlookpst

import (
	"fmt"
	"os"
	"time"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/ltp"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// Create creates a new PST file with the specified format.
// The file is created with a basic structure including:
// - Message store
// - Root folder
// - Named property map
func Create(filename string, format disk.PSTFormat) (*PST, error) {
	return CreateWithOptions(filename, CreateOptions{
		Format:      format,
		CryptMethod: disk.CryptMethodPermute,
	})
}

// CreateOptions contains options for creating a new PST file.
type CreateOptions struct {
	Format      disk.PSTFormat
	CryptMethod disk.CryptMethod
	DisplayName string // Display name for the PST
}

// CreateWithOptions creates a new PST file with the specified options.
func CreateWithOptions(filename string, opts CreateOptions) (*PST, error) {
	if opts.Format != disk.FormatUnicode && opts.Format != disk.FormatANSI {
		return nil, fmt.Errorf("invalid format: must be FormatUnicode or FormatANSI")
	}

	// Create the file
	f, err := os.Create(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to create file: %w", err)
	}

	// Initialize header
	header := disk.NewHeader(opts.Format, disk.ClientMagicPST)
	header.BCryptMethod = opts.CryptMethod

	// Create the initial file structure
	if err := initializePSTStructure(f, header, opts); err != nil {
		_ = f.Close()
		_ = os.Remove(filename)
		return nil, fmt.Errorf("failed to initialize PST structure: %w", err)
	}

	// Close and reopen in read-write mode
	_ = f.Close()

	return OpenReadWrite(filename)
}

// initializePSTStructure creates the initial PST file structure.
func initializePSTStructure(f *os.File, header *disk.Header, opts CreateOptions) error {
	format := header.Format

	// Create AMap manager for allocation
	amap := disk.CreateNewAMapManager(format)

	// Write initial header
	if err := disk.WriteHeader(f, header); err != nil {
		return fmt.Errorf("failed to write initial header: %w", err)
	}

	// Allocate and write DList page (at 0x4200)
	dlistData, err := disk.BuildDListPage(nil, 0, disk.DListPageLocation, format)
	if err != nil {
		return fmt.Errorf("failed to build DList page: %w", err)
	}
	if _, err := f.WriteAt(dlistData, disk.DListPageLocation); err != nil {
		return fmt.Errorf("failed to write DList page: %w", err)
	}

	// Write initial AMap page (at 0x4400)
	amapPage := amap.PendingAllocations() // Get initial AMap state
	_ = amapPage                          // AMap written through manager

	// Create minimal B-tree structure
	// We need at least:
	// 1. Message Store node (NID 0x21)
	// 2. Root Folder node (NID 0x2442)
	// 3. Name-to-ID map node (NID 0x61)

	// Allocate space for initial nodes
	nextBID := uint64(4) // Start after reserved BIDs

	// Create message store property context
	storeBag, err := ltp.CreateFolderPropertyBag(format, opts.DisplayName)
	if err != nil {
		return fmt.Errorf("failed to create store properties: %w", err)
	}
	if err := storeBag.SetString(ltp.PidTagMessageClass, "IPM.Note"); err != nil {
		return fmt.Errorf("failed to set message class: %w", err)
	}
	storeData, err := storeBag.Build()
	if err != nil {
		return fmt.Errorf("failed to build store properties: %w", err)
	}

	// Create root folder property context
	rootBag, err := ltp.CreateFolderPropertyBag(format, "Root")
	if err != nil {
		return fmt.Errorf("failed to create root folder properties: %w", err)
	}
	rootData, err := rootBag.Build()
	if err != nil {
		return fmt.Errorf("failed to build root folder properties: %w", err)
	}

	// Allocate blocks for nodes
	storeBlockOffset, err := amap.Allocate(disk.CalculateBlockDiskSize(uint64(len(storeData)), format))
	if err != nil {
		return fmt.Errorf("failed to allocate store block: %w", err)
	}
	storeBID := nextBID
	nextBID += 4

	rootBlockOffset, err := amap.Allocate(disk.CalculateBlockDiskSize(uint64(len(rootData)), format))
	if err != nil {
		return fmt.Errorf("failed to allocate root block: %w", err)
	}
	rootBID := nextBID
	nextBID += 4

	// Build and write data blocks
	storeBlock, err := disk.BuildBlock(storeData, storeBID, storeBlockOffset, format, header.BCryptMethod)
	if err != nil {
		return fmt.Errorf("failed to build store block: %w", err)
	}
	if _, err := f.WriteAt(storeBlock, int64(storeBlockOffset)); err != nil { //nolint:gosec // G115: offset bounded by file size
		return fmt.Errorf("failed to write store block: %w", err)
	}

	rootBlock, err := disk.BuildBlock(rootData, rootBID, rootBlockOffset, format, header.BCryptMethod)
	if err != nil {
		return fmt.Errorf("failed to build root block: %w", err)
	}
	if _, err := f.WriteAt(rootBlock, int64(rootBlockOffset)); err != nil { //nolint:gosec // G115: offset bounded by file size
		return fmt.Errorf("failed to write root block: %w", err)
	}

	// Create NBT entries
	nbtEntries := []disk.NBTLeafEntry{
		{
			NID:       uint64(util.NIDMessageStore),
			DataBID:   storeBID,
			SubBID:    0,
			ParentNID: 0,
		},
		{
			NID:       uint64(util.NIDRootFolder),
			DataBID:   rootBID,
			SubBID:    0,
			ParentNID: 0,
		},
	}

	// Create BBT entries
	bbtEntries := []disk.BBTLeafEntry{
		{
			BRef:     disk.BlockReference{BID: storeBID, IB: storeBlockOffset},
			Size:     uint16(len(storeData)), //nolint:gosec // G115: store data size bounded
			RefCount: 1,
		},
		{
			BRef:     disk.BlockReference{BID: rootBID, IB: rootBlockOffset},
			Size:     uint16(len(rootData)), //nolint:gosec // G115: root data size bounded
			RefCount: 1,
		},
	}

	// Allocate and write NBT root page
	nbtPageOffset, err := amap.Allocate(disk.PageSize)
	if err != nil {
		return fmt.Errorf("failed to allocate NBT page: %w", err)
	}
	nbtPage := disk.NewNBTLeafPage(nbtEntries, nbtPageOffset, format)
	nbtPageData, err := disk.BuildBTPage(nbtPage, format)
	if err != nil {
		return fmt.Errorf("failed to build NBT page: %w", err)
	}
	if _, err := f.WriteAt(nbtPageData, int64(nbtPageOffset)); err != nil { //nolint:gosec // G115: offset bounded by file size
		return fmt.Errorf("failed to write NBT page: %w", err)
	}

	// Allocate and write BBT root page
	bbtPageOffset, err := amap.Allocate(disk.PageSize)
	if err != nil {
		return fmt.Errorf("failed to allocate BBT page: %w", err)
	}
	bbtPage := disk.NewBBTLeafPage(bbtEntries, bbtPageOffset, format)
	bbtPageData, err := disk.BuildBTPage(bbtPage, format)
	if err != nil {
		return fmt.Errorf("failed to build BBT page: %w", err)
	}
	if _, err := f.WriteAt(bbtPageData, int64(bbtPageOffset)); err != nil { //nolint:gosec // G115: offset bounded by file size
		return fmt.Errorf("failed to write BBT page: %w", err)
	}

	// Persist the final allocation map after all initial blocks/pages have
	// been allocated. Reopening the PST must see those regions as occupied.
	if err := amap.WritePages(f); err != nil {
		return fmt.Errorf("failed to write allocation map: %w", err)
	}

	// Update header with B-tree roots
	header.Root.BRefNBT = disk.BlockReference{BID: nbtPageOffset, IB: nbtPageOffset}
	header.Root.BRefBBT = disk.BlockReference{BID: bbtPageOffset, IB: bbtPageOffset}
	header.Root.IBFileEOF = amap.FileSize()
	header.Root.CBAMapFree = amap.FreeSpace()
	header.BidNextB = nextBID
	header.SetAMapStatus(disk.AMapStatusValid)

	// Write final header
	if err := disk.WriteHeader(f, header); err != nil {
		return fmt.Errorf("failed to write final header: %w", err)
	}

	// Sync to disk
	return f.Sync()
}

// OpenReadWrite opens a PST file for reading and writing.
func OpenReadWrite(filename string) (*PST, error) {
	db, err := ndb.OpenReadWrite(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open PST: %w", err)
	}

	// Check for and perform recovery if needed
	result, err := ndb.CheckAndRecover(db)
	if err != nil {
		_ = db.Close()
		return nil, fmt.Errorf("recovery failed: %w", err)
	}
	if result.WasCorrupted {
		// Log recovery occurred (could add proper logging)
		_ = result
	}

	return &PST{db: db}, nil
}

// Save saves any pending changes to the PST file.
// If no transaction is active, this is a no-op.
func (p *PST) Save() error {
	if !p.db.IsDirty() {
		return nil
	}
	return p.db.Sync()
}

// BeginWrite starts a write transaction.
// All modifications must be made within a transaction.
func (p *PST) BeginWrite() (*WriteContext, error) {
	txn, err := p.db.BeginWrite()
	if err != nil {
		return nil, err
	}
	return &WriteContext{
		pst:          p,
		txn:          txn,
		tableWriters: make(map[util.NodeID]*ltp.TableWriter),
		dirtyTables:  make(map[util.NodeID]struct{}),
	}, nil
}

// WriteContext represents an active write transaction on a PST file.
type WriteContext struct {
	pst          *PST
	txn          *ndb.WriteTransaction
	tableWriters map[util.NodeID]*ltp.TableWriter
	dirtyTables  map[util.NodeID]struct{}
}

func (w *WriteContext) markTableDirty(nid util.NodeID) {
	if w.dirtyTables == nil {
		w.dirtyTables = make(map[util.NodeID]struct{})
	}
	w.dirtyTables[nid] = struct{}{}
}

func (w *WriteContext) flushDirtyTables() error {
	for nid := range w.dirtyTables {
		writer := w.tableWriters[nid]
		if writer == nil {
			continue
		}

		subnodes := ndb.NewSubnodeBuilder(w.txn)
		nextSubnodeIndex := uint32(1)
		data, err := writer.BuildWithSubnodes(subnodes, &nextSubnodeIndex)
		if err != nil {
			return fmt.Errorf("failed to build table 0x%X: %w", nid, err)
		}
		subBID, err := subnodes.Build()
		if err != nil {
			return fmt.Errorf("failed to build table 0x%X subnodes: %w", nid, err)
		}
		if err := ndb.UpdateNodeDataAndSubnodes(w.txn, nid, data, subBID); err != nil {
			return fmt.Errorf("failed to update table 0x%X: %w", nid, err)
		}
	}
	return nil
}

// Commit flushes logical table changes and then commits the transaction.
func (w *WriteContext) Commit() error {
	if err := w.flushDirtyTables(); err != nil {
		return err
	}
	return w.txn.Commit()
}

// Rollback discards all changes in the transaction.
func (w *WriteContext) Rollback() error {
	return w.txn.Rollback()
}

// Transaction returns the underlying NDB transaction.
func (w *WriteContext) Transaction() *ndb.WriteTransaction {
	return w.txn
}

// PST returns the parent PST file.
func (w *WriteContext) PST() *PST {
	return w.pst
}

// CreateFolder creates a new folder in the PST.
func (w *WriteContext) CreateFolder(parent *Folder, name string) (*Folder, error) {
	return CreateFolder(w, parent, name)
}

// CreateMessage creates a new message in a folder.
func (w *WriteContext) CreateMessage(folder *Folder) *MessageBuilder {
	return NewMessageBuilder(w, folder)
}

// DeleteFolder deletes a folder and all its contents.
func (w *WriteContext) DeleteFolder(folder *Folder) error {
	return DeleteFolder(w, folder)
}

// DeleteMessage deletes a message from a folder.
func (w *WriteContext) DeleteMessage(msg *Message) error {
	return DeleteMessage(w, msg)
}

// IsReadOnly returns true if the PST was opened in read-only mode.
func (p *PST) IsReadOnly() bool {
	return p.db.IsReadOnly()
}

// RecoverableInfo contains information about PST recovery.
type RecoverableInfo struct {
	NeedsRecovery  bool
	Status         disk.AMapStatus
	RecoveredCount int
}

// CheckRecovery checks if the PST needs recovery without performing it.
func CheckRecovery(filename string) (*RecoverableInfo, error) {
	// Open read-only to check status
	db, err := ndb.Open(filename)
	if err != nil {
		return nil, err
	}
	defer func() { _ = db.Close() }()

	status := db.Header().GetAMapStatus()
	return &RecoverableInfo{
		NeedsRecovery: status != disk.AMapStatusValid,
		Status:        status,
	}, nil
}

// CompactOptions contains options for compacting a PST file.
type CompactOptions struct {
	// RemoveDeletedItems removes items from deleted items folder
	RemoveDeletedItems bool
	// DefragmentBlocks consolidates fragmented blocks
	DefragmentBlocks bool
}

// Compact creates a compacted copy of the PST file.
// This removes unused space and defragments the file.
func Compact(srcFilename, dstFilename string, opts CompactOptions) error {
	// Open source PST
	src, err := Open(srcFilename)
	if err != nil {
		return fmt.Errorf("failed to open source: %w", err)
	}
	defer func() { _ = src.Close() }()

	// Create destination PST
	dst, err := Create(dstFilename, src.Format())
	if err != nil {
		return fmt.Errorf("failed to create destination: %w", err)
	}
	defer func() { _ = dst.Close() }()

	// Copy folder structure and messages
	// This is a simplified implementation
	srcRoot, err := src.RootFolder()
	if err != nil {
		return fmt.Errorf("failed to get source root: %w", err)
	}

	dstRoot, err := dst.RootFolder()
	if err != nil {
		return fmt.Errorf("failed to get destination root: %w", err)
	}

	// Begin write transaction
	ctx, err := dst.BeginWrite()
	if err != nil {
		return fmt.Errorf("failed to begin write: %w", err)
	}

	// Copy all subfolders recursively
	if err := copyFolderContents(ctx, srcRoot, dstRoot, opts); err != nil {
		_ = ctx.Rollback()
		return fmt.Errorf("failed to copy contents: %w", err)
	}

	// Commit
	if err := ctx.Commit(); err != nil {
		return fmt.Errorf("failed to commit: %w", err)
	}

	return nil
}

// copyFolderContents recursively copies folder contents.
// This is a best-effort operation that continues on individual item errors.
//
//nolint:unparam // Error return reserved for future use; currently always nil for best-effort semantics
func copyFolderContents(ctx *WriteContext, src, dst *Folder, opts CompactOptions) error {
	// Skip deleted items if requested
	if opts.RemoveDeletedItems {
		name, _ := src.Name()
		if name == "Deleted Items" {
			return nil
		}
	}

	// Copy messages
	for msg, err := range src.Messages() {
		if err != nil {
			continue
		}

		// Copy message
		subject, _ := msg.Subject()
		body, _ := msg.Body()
		sentTime, _ := msg.SubmitTime()

		builder := ctx.CreateMessage(dst)
		builder.SetSubject(subject)
		builder.SetBody(body)
		builder.SetSentTime(sentTime)

		if _, err := builder.Build(); err != nil {
			// Log but continue
			_ = err
		}
	}

	// Copy subfolders
	for subfolder, err := range src.Subfolders() {
		if err != nil {
			continue
		}

		name, _ := subfolder.Name()

		// Create subfolder in destination
		dstSubfolder, err := ctx.CreateFolder(dst, name)
		if err != nil {
			continue
		}

		// Recursively copy contents
		if err := copyFolderContents(ctx, subfolder, dstSubfolder, opts); err != nil {
			// Log but continue
			_ = err
		}
	}

	return nil
}

// Statistics contains PST file statistics.
type Statistics struct {
	Format          disk.PSTFormat
	FileSize        uint64
	FreeSpace       uint64
	FolderCount     int
	MessageCount    int
	AttachmentCount int
}

// GetStatistics returns statistics about the PST file.
func (p *PST) GetStatistics() (*Statistics, error) {
	stats := &Statistics{
		Format:    p.Format(),
		FileSize:  p.db.Header().Root.IBFileEOF,
		FreeSpace: p.db.Header().Root.CBAMapFree,
	}

	// Count folders and messages
	root, err := p.RootFolder()
	if err != nil {
		return stats, nil
	}

	countFolderContents(root, stats)
	return stats, nil
}

// countFolderContents recursively counts folders and messages.
func countFolderContents(folder *Folder, stats *Statistics) {
	stats.FolderCount++

	msgCount, _ := folder.ContentCount()
	stats.MessageCount += int(msgCount)

	for subfolder, err := range folder.Subfolders() {
		if err != nil {
			continue
		}
		countFolderContents(subfolder, stats)
	}
}

// ExportOptions contains options for exporting messages.
type ExportOptions struct {
	IncludeAttachments bool
	DateRange          *DateRange
}

// DateRange specifies a date range for filtering.
type DateRange struct {
	Start time.Time
	End   time.Time
}
