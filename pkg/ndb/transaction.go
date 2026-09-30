package ndb

import (
	"errors"
	"fmt"
	"sync"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// TransactionState represents the current state of a write transaction.
type TransactionState int

const (
	// TransactionStateIdle indicates no active transaction.
	TransactionStateIdle TransactionState = iota
	// TransactionStateActive indicates a transaction is in progress.
	TransactionStateActive
	// TransactionStateCommitting indicates commit phase 1 is in progress.
	TransactionStateCommitting
	// TransactionStateCompleting indicates commit phase 2 is in progress.
	TransactionStateCompleting
)

// WriteTransaction represents an atomic write operation on a PST database.
// It implements a two-phase commit protocol following the Microsoft Rust PST SDK pattern:
//
//	Phase 1 (startWrite): Set AMapStatus = Invalid, write all data
//	Phase 2 (finishWrite): Set AMapStatus = Valid2, update header
//
// If a crash occurs between phases, the PST file can be recovered.
type WriteTransaction struct {
	db       *Database
	amap     *disk.AMapManager
	btwriter *BTWriter

	state TransactionState
	mu    sync.Mutex

	// Pending writes
	pendingBlocks []pendingBlock
	pendingNodes  []pendingNode

	// New B-tree roots after commit
	newNBTRoot *disk.BlockReference
	newBBTRoot *disk.BlockReference

	// Original header values for rollback
	originalHeader *disk.Header
}

// pendingBlock represents a block to be written.
type pendingBlock struct {
	bid    util.BlockID
	offset uint64
	data   []byte
	size   uint16
}

// pendingNode represents a node to be created/modified.
type pendingNode struct {
	info *NodeInfo
}

// BeginWrite starts a new write transaction on the database.
// Only one transaction can be active at a time.
func (db *Database) BeginWrite() (*WriteTransaction, error) {
	if db.readOnly {
		return nil, errors.New("database is read-only")
	}

	// Create AMap manager if needed
	amap, err := disk.NewAMapManager(db.file, db.header)
	if err != nil {
		return nil, fmt.Errorf("failed to create AMap manager: %w", err)
	}

	txn := &WriteTransaction{
		db:       db,
		amap:     amap,
		btwriter: NewBTWriter(db, amap),
		state:    TransactionStateActive,
	}

	// Save original header for potential rollback
	txn.originalHeader = cloneHeader(db.header)

	return txn, nil
}

// cloneHeader creates a deep copy of the header.
func cloneHeader(h *disk.Header) *disk.Header {
	if h == nil {
		return nil
	}
	clone := *h
	return &clone
}

// WriteBlockData writes data as a new block and returns its BID.
// The data is queued and written during commit.
func (t *WriteTransaction) WriteBlockData(data []byte) (util.BlockID, error) {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return 0, errors.New("transaction is not active")
	}

	// Mark database as dirty when modifications are made
	t.db.setDirty()

	// Check data size
	maxSize := disk.MaxDataBlockSizeUnicode
	if t.db.Format() == disk.FormatANSI {
		maxSize = disk.MaxDataBlockSizeANSI
	}
	if len(data) > maxSize {
		return 0, fmt.Errorf("data too large for single block: %d bytes (max %d)", len(data), maxSize)
	}

	// Allocate block ID
	bid := t.db.AllocateBlockID()

	// Calculate disk size and allocate space
	diskSize := disk.CalculateBlockDiskSize(uint64(len(data)), t.db.Format())
	offset, err := t.amap.Allocate(diskSize)
	if err != nil {
		return 0, fmt.Errorf("failed to allocate space: %w", err)
	}

	// Queue for writing
	t.pendingBlocks = append(t.pendingBlocks, pendingBlock{
		bid:    bid,
		offset: offset,
		data:   data,
		size:   uint16(len(data)), //nolint:gosec // G115: block size bounded by MaxBlockSize
	})

	// Queue BBT entry
	if err := t.btwriter.InsertBlock(&BlockInfo{
		BID:      bid,
		Location: offset,
		Size:     uint16(len(data)), //nolint:gosec // G115: block size bounded by MaxBlockSize
		RefCount: 1,
	}); err != nil {
		return 0, err
	}

	return bid, nil
}

// WriteExtendedBlockData writes data that may span multiple blocks.
// For data larger than max block size, creates XBLOCK/XXBLOCK structures.
func (t *WriteTransaction) WriteExtendedBlockData(data []byte) (util.BlockID, error) {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return 0, errors.New("transaction is not active")
	}

	maxSize := disk.MaxDataBlockSizeUnicode
	if t.db.Format() == disk.FormatANSI {
		maxSize = disk.MaxDataBlockSizeANSI
	}

	// If data fits in single block, use simple write
	if len(data) <= maxSize {
		t.mu.Unlock()
		bid, err := t.WriteBlockData(data)
		t.mu.Lock()
		return bid, err
	}

	// Split data into chunks
	var dataBlockBIDs []uint64
	for offset := 0; offset < len(data); offset += maxSize {
		end := offset + maxSize
		if end > len(data) {
			end = len(data)
		}
		chunk := data[offset:end]

		// Write chunk as separate block
		t.mu.Unlock()
		chunkBID, err := t.WriteBlockData(chunk)
		t.mu.Lock()
		if err != nil {
			return 0, err
		}
		dataBlockBIDs = append(dataBlockBIDs, uint64(chunkBID))
	}

	// Create XBLOCK pointing to data blocks
	xblockBID := t.db.AllocateInternalBlockID()
	xblockDiskSize := disk.CalculateBlockDiskSize(uint64(8+len(dataBlockBIDs)*8), t.db.Format())
	xblockOffset, err := t.amap.Allocate(xblockDiskSize)
	if err != nil {
		return 0, err
	}

	xblockData, err := disk.BuildExtendedBlock(1, dataBlockBIDs, uint32(len(data)), uint64(xblockBID), xblockOffset, t.db.Format()) //nolint:gosec // G115: len(data) bounded by caller
	if err != nil {
		return 0, err
	}

	t.pendingBlocks = append(t.pendingBlocks, pendingBlock{
		bid:    xblockBID,
		offset: xblockOffset,
		data:   xblockData,
		size:   uint16(len(xblockData)), //nolint:gosec // G115: xblock size bounded
	})

	if err := t.btwriter.InsertBlock(&BlockInfo{
		BID:      xblockBID,
		Location: xblockOffset,
		Size:     uint16(8 + len(dataBlockBIDs)*8), //nolint:gosec // G115: xblock header size bounded
		RefCount: 1,
	}); err != nil {
		return 0, err
	}

	return xblockBID, nil
}

// AllocateNodeID reserves a new node ID for the requested node type.
func (t *WriteTransaction) AllocateNodeID(nidType util.NIDType) (util.NodeID, error) {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return 0, errors.New("transaction is not active")
	}

	t.db.mu.Lock()
	t.db.nextNID++
	index := t.db.nextNID
	t.db.mu.Unlock()

	return util.MakeNID(nidType, index), nil
}

// CreateNode creates a new node in the NBT.
func (t *WriteTransaction) CreateNode(nidType util.NIDType, parentNID util.NodeID, dataBID, subBID util.BlockID) (*NodeInfo, error) {
	nid, err := t.AllocateNodeID(nidType)
	if err != nil {
		return nil, err
	}
	return t.CreateNodeWithID(nid, parentNID, dataBID, subBID)
}

// CreateNodeWithID creates a new node using an already selected NID.
// This is required for PST table nodes whose nidIndex must match their owner.
func (t *WriteTransaction) CreateNodeWithID(nid util.NodeID, parentNID util.NodeID, dataBID, subBID util.BlockID) (*NodeInfo, error) {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return nil, errors.New("transaction is not active")
	}

	info := &NodeInfo{
		NID:       nid,
		DataBID:   dataBID,
		SubBID:    subBID,
		ParentNID: parentNID,
	}

	if err := t.btwriter.InsertNode(info); err != nil {
		return nil, err
	}

	t.pendingNodes = append(t.pendingNodes, pendingNode{info: info})
	return info, nil
}

// DeleteNode marks a node for deletion.
func (t *WriteTransaction) DeleteNode(nid util.NodeID) error {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return errors.New("transaction is not active")
	}

	return t.btwriter.DeleteNode(nid)
}

// DeleteBlock marks a block for deletion.
func (t *WriteTransaction) DeleteBlock(bid util.BlockID) error {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return errors.New("transaction is not active")
	}

	// Look up block info to free its space
	info, err := t.db.LookupBlock(bid)
	if err == nil {
		diskSize := disk.CalculateBlockDiskSize(uint64(info.Size), t.db.Format())
		if err := t.amap.Free(info.Location, diskSize); err != nil {
			// Log but don't fail - space leak is acceptable
			_ = err
		}
	}

	return t.btwriter.DeleteBlock(bid)
}

// Commit commits the transaction, writing all changes to disk.
// This implements a two-phase commit for crash safety.
func (t *WriteTransaction) Commit() error {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state != TransactionStateActive {
		return errors.New("transaction is not active")
	}

	// Phase 1: Start write - mark AMap as invalid
	if err := t.startWrite(); err != nil {
		return fmt.Errorf("commit phase 1 failed: %w", err)
	}

	t.state = TransactionStateCommitting

	// Write all pending blocks
	for _, pb := range t.pendingBlocks {
		blockData, err := disk.BuildBlock(pb.data, uint64(pb.bid), pb.offset, t.db.Format(), t.db.CryptMethod())
		if err != nil {
			return fmt.Errorf("failed to build block 0x%X: %w", pb.bid, err)
		}
		if err := t.db.WriteBlock(pb.offset, blockData); err != nil {
			return fmt.Errorf("failed to write block 0x%X: %w", pb.bid, err)
		}
	}

	// Apply B-tree changes
	nbtRoot, bbtRoot, err := t.btwriter.Apply()
	if err != nil {
		return fmt.Errorf("failed to apply B-tree changes: %w", err)
	}
	t.newNBTRoot = nbtRoot
	t.newBBTRoot = bbtRoot

	// Persist allocation changes while the header is still marked invalid.
	// A later transaction must never see live blocks as free.
	if err := t.amap.WritePages(t.db.file); err != nil {
		return fmt.Errorf("failed to write allocation maps: %w", err)
	}

	// Sync to ensure data and allocation maps are on disk before phase 2.
	if err := t.db.Sync(); err != nil {
		return fmt.Errorf("sync failed: %w", err)
	}

	t.state = TransactionStateCompleting

	// Phase 2: Finish write - update header with new roots and mark valid
	if err := t.finishWrite(); err != nil {
		return fmt.Errorf("commit phase 2 failed: %w", err)
	}

	// Final sync
	if err := t.db.Sync(); err != nil {
		return fmt.Errorf("final sync failed: %w", err)
	}

	// Clear dirty flag and invalidate caches
	t.db.clearDirty()
	t.db.InvalidateCache()
	t.db.ReloadBTrees()

	t.state = TransactionStateIdle

	return nil
}

// startWrite begins the two-phase commit by marking the AMap as invalid.
func (t *WriteTransaction) startWrite() error {
	t.db.header.SetAMapStatus(disk.AMapStatusInvalid)
	t.db.header.IncrementUnique()

	// Write just the header with invalid status
	return disk.WriteHeader(t.db.file, t.db.header)
}

// finishWrite completes the two-phase commit by updating roots and marking valid.
func (t *WriteTransaction) finishWrite() error {
	// Update B-tree roots
	if t.newNBTRoot != nil {
		t.db.header.Root.BRefNBT = *t.newNBTRoot
	}
	if t.newBBTRoot != nil {
		t.db.header.Root.BRefBBT = *t.newBBTRoot
	}

	// Update file size
	t.db.header.UpdateFileSize(t.amap.FileSize())

	// Update next block ID
	t.db.header.UpdateNextBlockID(uint64(t.db.NextBlockID()))

	// Update AMap info
	t.db.header.Root.CBAMapFree = t.amap.FreeSpace()

	// Mark as valid
	t.db.header.SetAMapStatus(disk.AMapStatusValid2)

	// Write updated header
	return disk.WriteHeader(t.db.file, t.db.header)
}

// Rollback discards all pending changes.
func (t *WriteTransaction) Rollback() error {
	t.mu.Lock()
	defer t.mu.Unlock()

	if t.state == TransactionStateIdle {
		return nil // Nothing to rollback
	}

	// Restore original header if we started writing
	if t.state >= TransactionStateCommitting && t.originalHeader != nil {
		*t.db.header = *t.originalHeader
		if err := disk.WriteHeader(t.db.file, t.db.header); err != nil {
			return fmt.Errorf("failed to restore header: %w", err)
		}
	}

	// Clear pending changes
	t.btwriter.Reset()
	t.pendingBlocks = nil
	t.pendingNodes = nil
	t.newNBTRoot = nil
	t.newBBTRoot = nil

	t.state = TransactionStateIdle

	return nil
}

// State returns the current transaction state.
func (t *WriteTransaction) State() TransactionState {
	t.mu.Lock()
	defer t.mu.Unlock()
	return t.state
}

// Database returns the underlying database.
func (t *WriteTransaction) Database() *Database {
	return t.db
}

// AMapManager returns the allocation map manager.
func (t *WriteTransaction) AMapManager() *disk.AMapManager {
	return t.amap
}
