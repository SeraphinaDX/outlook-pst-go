package ndb

import (
	"fmt"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// NodeBuilder provides a fluent interface for creating nodes.
type NodeBuilder struct {
	txn       *WriteTransaction
	nidType   util.NIDType
	nid       util.NodeID
	parentNID util.NodeID
	dataBID   util.BlockID
	subBID    util.BlockID
	data      []byte
}

// NewNodeBuilder creates a new node builder.
func NewNodeBuilder(txn *WriteTransaction, nidType util.NIDType) *NodeBuilder {
	return &NodeBuilder{
		txn:     txn,
		nidType: nidType,
	}
}

// WithNID sets an explicit node ID. Table nodes use the same nidIndex as their owner.
func (b *NodeBuilder) WithNID(nid util.NodeID) *NodeBuilder {
	b.nid = nid
	return b
}

// WithParent sets the parent node ID.
func (b *NodeBuilder) WithParent(parentNID util.NodeID) *NodeBuilder {
	b.parentNID = parentNID
	return b
}

// WithData sets the node data. Data will be written as a block during build.
func (b *NodeBuilder) WithData(data []byte) *NodeBuilder {
	b.data = data
	return b
}

// WithDataBID sets a pre-existing data block ID.
func (b *NodeBuilder) WithDataBID(bid util.BlockID) *NodeBuilder {
	b.dataBID = bid
	return b
}

// WithSubnodeBID sets a pre-existing subnode block ID.
func (b *NodeBuilder) WithSubnodeBID(bid util.BlockID) *NodeBuilder {
	b.subBID = bid
	return b
}

// Build creates the node and returns its info.
func (b *NodeBuilder) Build() (*NodeInfo, error) {
	// Write data if provided
	if len(b.data) > 0 {
		maxSize := disk.MaxDataBlockSizeUnicode
		if b.txn.db.Format() == disk.FormatANSI {
			maxSize = disk.MaxDataBlockSizeANSI
		}

		var err error
		if len(b.data) <= maxSize {
			b.dataBID, err = b.txn.WriteBlockData(b.data)
		} else {
			b.dataBID, err = b.txn.WriteExtendedBlockData(b.data)
		}
		if err != nil {
			return nil, fmt.Errorf("failed to write node data: %w", err)
		}
	}

	// Create the node
	if b.nid != 0 {
		if b.nid.Type() != b.nidType {
			return nil, fmt.Errorf("explicit NID type 0x%X does not match builder type 0x%X", b.nid.Type(), b.nidType)
		}
		return b.txn.CreateNodeWithID(b.nid, b.parentNID, b.dataBID, b.subBID)
	}
	return b.txn.CreateNode(b.nidType, b.parentNID, b.dataBID, b.subBID)
}

// SubnodeBuilder provides a fluent interface for building subnode trees.
type SubnodeBuilder struct {
	txn     *WriteTransaction
	entries []disk.SubnodeLeafEntry
}

// NewSubnodeBuilder creates a new subnode builder.
func NewSubnodeBuilder(txn *WriteTransaction) *SubnodeBuilder {
	return &SubnodeBuilder{
		txn: txn,
	}
}

// AddSubnode adds a subnode entry.
func (b *SubnodeBuilder) AddSubnode(nid util.NodeID, data []byte) error {
	return b.AddSubnodeWithSubnodes(nid, data, 0)
}

// AddSubnodeWithSubnodes adds a subnode entry whose node may itself own a
// subnode tree. This is used for objects such as attachments whose large
// properties are stored indirectly beneath the attachment node.
func (b *SubnodeBuilder) AddSubnodeWithSubnodes(nid util.NodeID, data []byte, subBID util.BlockID) error {
	var dataBID util.BlockID
	var err error

	if len(data) > 0 {
		maxSize := disk.MaxDataBlockSizeUnicode
		if b.txn.db.Format() == disk.FormatANSI {
			maxSize = disk.MaxDataBlockSizeANSI
		}

		if len(data) <= maxSize {
			dataBID, err = b.txn.WriteBlockData(data)
		} else {
			dataBID, err = b.txn.WriteExtendedBlockData(data)
		}
		if err != nil {
			return fmt.Errorf("failed to write subnode data: %w", err)
		}
	}

	b.entries = append(b.entries, disk.SubnodeLeafEntry{
		NID:     uint64(nid),
		DataBID: uint64(dataBID),
		SubBID:  uint64(subBID),
	})

	return nil
}

// AddSubnodeWithBID adds a subnode entry with a pre-existing data BID.
func (b *SubnodeBuilder) AddSubnodeWithBID(nid util.NodeID, dataBID util.BlockID) {
	b.entries = append(b.entries, disk.SubnodeLeafEntry{
		NID:     uint64(nid),
		DataBID: uint64(dataBID),
		SubBID:  0,
	})
}

// Build creates the subnode block and returns its BID.
func (b *SubnodeBuilder) Build() (util.BlockID, error) {
	if len(b.entries) == 0 {
		return 0, nil
	}

	// Allocate subnode block ID (internal)
	bid := b.txn.db.AllocateInternalBlockID()

	// Calculate disk size
	format := b.txn.db.Format()
	entrySize := 24 // Unicode
	headerSize := 8
	if format == disk.FormatANSI {
		entrySize = 12
		headerSize = 4
	}
	dataSize := uint64(headerSize + len(b.entries)*entrySize)
	diskSize := disk.CalculateBlockDiskSize(dataSize, format)

	// Allocate space
	offset, err := b.txn.amap.Allocate(diskSize)
	if err != nil {
		return 0, fmt.Errorf("failed to allocate subnode block: %w", err)
	}

	// Build subnode block
	blockData, err := disk.BuildSubnodeLeafBlock(b.entries, uint64(bid), offset, format)
	if err != nil {
		return 0, fmt.Errorf("failed to build subnode block: %w", err)
	}

	// Queue for writing
	b.txn.pendingBlocks = append(b.txn.pendingBlocks, pendingBlock{
		bid:    bid,
		offset: offset,
		data:   blockData,
		size:   uint16(len(blockData)), //nolint:gosec // G115: block size bounded by MaxBlockSize
	})

	// Register in BBT
	if err := b.txn.btwriter.InsertBlock(&BlockInfo{
		BID:      bid,
		Location: offset,
		Size:     uint16(dataSize), //nolint:gosec // G115: data size bounded by MaxDataSize
		RefCount: 1,
	}); err != nil {
		return 0, err
	}

	return bid, nil
}

// UpdateNodeData updates the data block for an existing node.
// This creates a new data block and updates the NBT entry.
func UpdateNodeData(txn *WriteTransaction, nid util.NodeID, data []byte) error {
	// Newly-created nodes are not visible through the on-disk NBT until commit.
	// Prefer a pending insert when present so callers can build a folder and add
	// messages to it in the same transaction.
	info, pending := txn.btwriter.nbtInserts[nid]
	var err error
	if !pending {
		info, err = txn.db.LookupNode(nid)
		if err != nil {
			return fmt.Errorf("node not found: %w", err)
		}
	}

	// Delete an old committed data block when replacing an existing node.
	if !pending && info.DataBID != 0 {
		if err := txn.DeleteBlock(info.DataBID); err != nil {
			_ = err
		}
	}

	// Write new data
	var newBID util.BlockID
	if len(data) > 0 {
		maxSize := disk.MaxDataBlockSizeUnicode
		if txn.db.Format() == disk.FormatANSI {
			maxSize = disk.MaxDataBlockSizeANSI
		}

		if len(data) <= maxSize {
			newBID, err = txn.WriteBlockData(data)
		} else {
			newBID, err = txn.WriteExtendedBlockData(data)
		}
		if err != nil {
			return fmt.Errorf("failed to write new data: %w", err)
		}
	}

	if pending {
		// nbtInserts stores the same *NodeInfo pointer, so updating it in place
		// changes the queued NBT entry without attempting a duplicate insert.
		info.DataBID = newBID
		return nil
	}

	if err := txn.btwriter.DeleteNode(nid); err != nil {
		return err
	}

	newInfo := &NodeInfo{
		NID:       nid,
		DataBID:   newBID,
		SubBID:    info.SubBID,
		ParentNID: info.ParentNID,
	}
	return txn.btwriter.InsertNode(newInfo)
}

// UpdateNodeSubnodes updates the subnode block for an existing node.
func UpdateNodeSubnodes(txn *WriteTransaction, nid util.NodeID, subBID util.BlockID) error {
	info, err := txn.db.LookupNode(nid)
	if err != nil {
		return fmt.Errorf("node not found: %w", err)
	}

	// Delete old subnode block if present
	if info.SubBID != 0 {
		if err := txn.DeleteBlock(info.SubBID); err != nil {
			_ = err
		}
	}

	// Delete old NBT entry and insert updated one
	if err := txn.btwriter.DeleteNode(nid); err != nil {
		return err
	}

	newInfo := &NodeInfo{
		NID:       nid,
		DataBID:   info.DataBID,
		SubBID:    subBID,
		ParentNID: info.ParentNID,
	}

	return txn.btwriter.InsertNode(newInfo)
}

// CopyNode creates a copy of an existing node with a new NID.
// The data and subnode blocks are referenced (not copied).
func CopyNode(txn *WriteTransaction, sourceNID util.NodeID, newType util.NIDType, newParentNID util.NodeID) (*NodeInfo, error) {
	// Look up source node
	sourceInfo, err := txn.db.LookupNode(sourceNID)
	if err != nil {
		return nil, fmt.Errorf("source node not found: %w", err)
	}

	// Create new node with same blocks
	return txn.CreateNode(newType, newParentNID, sourceInfo.DataBID, sourceInfo.SubBID)
}

// DeleteNodeRecursive deletes a node and all its subnodes.
func DeleteNodeRecursive(txn *WriteTransaction, nid util.NodeID) error {
	// Look up node
	info, err := txn.db.LookupNode(nid)
	if err != nil {
		// Node doesn't exist - nothing to delete
		return nil
	}

	// Delete subnodes if present
	if info.SubBID != 0 {
		node, err := txn.db.GetNode(nid)
		if err == nil && node != nil {
			// Iterate subnodes and delete them
			for subNode, subErr := range node.Subnodes() {
				if subErr != nil {
					continue
				}
				if err := DeleteNodeRecursive(txn, subNode.ID()); err != nil {
					// Log but continue
					_ = err
				}
			}
		}

		// Delete subnode block
		if err := txn.DeleteBlock(info.SubBID); err != nil {
			_ = err
		}
	}

	// Delete data block
	if info.DataBID != 0 {
		if err := txn.DeleteBlock(info.DataBID); err != nil {
			_ = err
		}
	}

	// Delete node from NBT
	return txn.DeleteNode(nid)
}

// MoveNode changes the parent of a node.
func MoveNode(txn *WriteTransaction, nid util.NodeID, newParentNID util.NodeID) error {
	info, err := txn.db.LookupNode(nid)
	if err != nil {
		return fmt.Errorf("node not found: %w", err)
	}

	// Delete old entry
	if err := txn.btwriter.DeleteNode(nid); err != nil {
		return err
	}

	// Insert with new parent
	newInfo := &NodeInfo{
		NID:       nid,
		DataBID:   info.DataBID,
		SubBID:    info.SubBID,
		ParentNID: newParentNID,
	}

	return txn.btwriter.InsertNode(newInfo)
}

// CreateFolderNode creates a new folder node structure.
// Folders require a hierarchy table and contents table in addition to properties.
func CreateFolderNode(txn *WriteTransaction, parentNID util.NodeID, folderData []byte) (*NodeInfo, error) {
	// Create main folder node
	folderNID, err := NewNodeBuilder(txn, util.NIDTypeNormalFolder).
		WithParent(parentNID).
		WithData(folderData).
		Build()
	if err != nil {
		return nil, fmt.Errorf("failed to create folder node: %w", err)
	}

	// Create hierarchy table node (for subfolders)
	// This would typically have empty table context data
	_, err = NewNodeBuilder(txn, util.NIDTypeHierarchyTable).
		WithParent(folderNID.NID).
		Build()
	if err != nil {
		return nil, fmt.Errorf("failed to create hierarchy table: %w", err)
	}

	// Create contents table node (for messages)
	_, err = NewNodeBuilder(txn, util.NIDTypeContentsTable).
		WithParent(folderNID.NID).
		Build()
	if err != nil {
		return nil, fmt.Errorf("failed to create contents table: %w", err)
	}

	return folderNID, nil
}

// CreateMessageNode creates a new message node structure.
func CreateMessageNode(txn *WriteTransaction, parentNID util.NodeID, messageData []byte) (*NodeInfo, error) {
	return NewNodeBuilder(txn, util.NIDTypeNormalMessage).
		WithParent(parentNID).
		WithData(messageData).
		Build()
}
