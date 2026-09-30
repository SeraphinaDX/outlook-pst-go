package ndb

import (
	"fmt"
	"sort"

	"github.com/grokify/outlook-pst-go/pkg/disk"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// BTWriter handles B-tree modifications for both NBT and BBT.
// It buffers changes and writes them during transaction commit.
type BTWriter struct {
	db   *Database
	amap *disk.AMapManager

	// Pending changes
	nbtInserts map[util.NodeID]*NodeInfo
	nbtDeletes map[util.NodeID]bool
	bbtInserts map[util.BlockID]*BlockInfo
	bbtDeletes map[util.BlockID]bool

	// Modified pages to write
	modifiedNBTPages []*disk.BTPage
	modifiedBBTPages []*disk.BTPage

	// New root pages (if tree structure changed)
	newNBTRoot *disk.BTPage
	newBBTRoot *disk.BTPage
}

// NewBTWriter creates a new B-tree writer.
func NewBTWriter(db *Database, amap *disk.AMapManager) *BTWriter {
	return &BTWriter{
		db:         db,
		amap:       amap,
		nbtInserts: make(map[util.NodeID]*NodeInfo),
		nbtDeletes: make(map[util.NodeID]bool),
		bbtInserts: make(map[util.BlockID]*BlockInfo),
		bbtDeletes: make(map[util.BlockID]bool),
	}
}

// InsertNode queues a node for insertion into the NBT.
func (w *BTWriter) InsertNode(info *NodeInfo) error {
	if info == nil {
		return fmt.Errorf("node info cannot be nil")
	}
	if _, exists := w.nbtInserts[info.NID]; exists {
		return fmt.Errorf("node 0x%X already queued for insertion", info.NID)
	}
	if w.nbtDeletes[info.NID] {
		// Reinserting a deleted node
		delete(w.nbtDeletes, info.NID)
	}
	w.nbtInserts[info.NID] = info
	return nil
}

// DeleteNode queues a node for deletion from the NBT.
func (w *BTWriter) DeleteNode(nid util.NodeID) error {
	if w.nbtDeletes[nid] {
		return fmt.Errorf("node 0x%X already queued for deletion", nid)
	}
	if _, exists := w.nbtInserts[nid]; exists {
		// Cancel pending insert
		delete(w.nbtInserts, nid)
		return nil
	}
	w.nbtDeletes[nid] = true
	return nil
}

// InsertBlock queues a block for insertion into the BBT.
func (w *BTWriter) InsertBlock(info *BlockInfo) error {
	if info == nil {
		return fmt.Errorf("block info cannot be nil")
	}
	lookupBID := info.BID &^ util.BlockID(0x2) // Clear internal bit for map key
	if _, exists := w.bbtInserts[lookupBID]; exists {
		return fmt.Errorf("block 0x%X already queued for insertion", info.BID)
	}
	if w.bbtDeletes[lookupBID] {
		delete(w.bbtDeletes, lookupBID)
	}
	w.bbtInserts[lookupBID] = info
	return nil
}

// DeleteBlock queues a block for deletion from the BBT.
func (w *BTWriter) DeleteBlock(bid util.BlockID) error {
	lookupBID := bid &^ util.BlockID(0x2)
	if w.bbtDeletes[lookupBID] {
		return fmt.Errorf("block 0x%X already queued for deletion", bid)
	}
	if _, exists := w.bbtInserts[lookupBID]; exists {
		delete(w.bbtInserts, lookupBID)
		return nil
	}
	w.bbtDeletes[lookupBID] = true
	return nil
}

// Apply applies all pending changes to the B-trees.
// This rebuilds affected pages and returns the new root references.
func (w *BTWriter) Apply() (nbtRoot, bbtRoot *disk.BlockReference, err error) {
	// Apply NBT changes
	nbtRef, err := w.applyNBTChanges()
	if err != nil {
		return nil, nil, fmt.Errorf("failed to apply NBT changes: %w", err)
	}

	// Apply BBT changes
	bbtRef, err := w.applyBBTChanges()
	if err != nil {
		return nil, nil, fmt.Errorf("failed to apply BBT changes: %w", err)
	}

	return nbtRef, bbtRef, nil
}

// applyNBTChanges applies pending NBT changes.
func (w *BTWriter) applyNBTChanges() (*disk.BlockReference, error) {
	if len(w.nbtInserts) == 0 && len(w.nbtDeletes) == 0 {
		// No changes - return existing root
		return &w.db.header.Root.BRefNBT, nil
	}

	// Load current NBT root
	if err := w.db.loadNBTRoot(); err != nil {
		return nil, err
	}

	// Collect all current entries
	entries, err := w.collectNBTEntries(w.db.nbtRoot)
	if err != nil {
		return nil, err
	}

	// Apply deletes
	filteredEntries := make([]disk.NBTLeafEntry, 0, len(entries))
	for _, entry := range entries {
		nid := util.NodeID(entry.NID) //nolint:gosec // G115: NID is 32-bit per MS-PST spec
		if !w.nbtDeletes[nid] {
			filteredEntries = append(filteredEntries, entry)
		}
	}

	// Apply inserts
	for _, info := range w.nbtInserts {
		entry := disk.NBTLeafEntry{
			NID:       uint64(info.NID),
			DataBID:   uint64(info.DataBID),
			SubBID:    uint64(info.SubBID),
			ParentNID: uint32(info.ParentNID),
		}
		filteredEntries = append(filteredEntries, entry)
	}

	// Sort by NID
	sort.Slice(filteredEntries, func(i, j int) bool {
		return filteredEntries[i].NID < filteredEntries[j].NID
	})

	// Rebuild tree structure
	return w.rebuildNBT(filteredEntries)
}

// collectNBTEntries recursively collects all leaf entries from the NBT.
func (w *BTWriter) collectNBTEntries(page *disk.BTPage) ([]disk.NBTLeafEntry, error) {
	if page.IsLeaf() {
		return page.NBTEntries, nil
	}

	// Non-leaf - recursively collect from children
	var entries []disk.NBTLeafEntry
	for _, child := range page.NonleafEntries {
		childData, err := w.db.readPage(child.Ref.IB)
		if err != nil {
			return nil, err
		}
		childPage, err := disk.ParseBTPage(childData, w.db.Format(), disk.PageTypeNBT)
		if err != nil {
			return nil, err
		}
		childEntries, err := w.collectNBTEntries(childPage)
		if err != nil {
			return nil, err
		}
		entries = append(entries, childEntries...)
	}
	return entries, nil
}

// rebuildNBT rebuilds the NBT from a sorted list of entries.
func (w *BTWriter) rebuildNBT(entries []disk.NBTLeafEntry) (*disk.BlockReference, error) {
	if len(entries) == 0 {
		// Empty tree - create minimal valid structure
		return w.createEmptyNBT()
	}

	format := w.db.Format()
	maxEntriesPerPage := disk.MaxBTPageEntriesNBTUnicode
	if format == disk.FormatANSI {
		maxEntriesPerPage = disk.MaxBTPageEntriesNBTANSI
	}

	// If all entries fit in one page, create single leaf page
	if len(entries) <= maxEntriesPerPage {
		return w.createNBTLeafPage(entries)
	}

	// Need multiple pages - create leaf pages and build tree
	return w.buildNBTTree(entries, maxEntriesPerPage)
}

// createEmptyNBT creates an empty NBT with a single leaf page.
func (w *BTWriter) createEmptyNBT() (*disk.BlockReference, error) {
	return w.createNBTLeafPage([]disk.NBTLeafEntry{})
}

// createNBTLeafPage creates a single NBT leaf page.
func (w *BTWriter) createNBTLeafPage(entries []disk.NBTLeafEntry) (*disk.BlockReference, error) {
	// Allocate space for the page
	offset, err := w.amap.Allocate(disk.PageSize)
	if err != nil {
		return nil, err
	}

	// Create page
	bid := offset // Use offset as BID for pages
	page := disk.NewNBTLeafPage(entries, bid, w.db.Format())

	// Serialize and write
	pageData, err := disk.BuildBTPage(page, w.db.Format())
	if err != nil {
		return nil, err
	}

	if err := w.db.WritePage(offset, pageData); err != nil {
		return nil, err
	}

	w.newNBTRoot = page

	return &disk.BlockReference{
		BID: bid,
		IB:  offset,
	}, nil
}

// buildNBTTree builds a multi-level NBT from entries.
func (w *BTWriter) buildNBTTree(entries []disk.NBTLeafEntry, maxPerPage int) (*disk.BlockReference, error) {
	// Split entries into leaf pages
	var leafRefs []disk.BlockReference
	var leafMinKeys []uint64

	for i := 0; i < len(entries); i += maxPerPage {
		end := i + maxPerPage
		if end > len(entries) {
			end = len(entries)
		}
		pageEntries := entries[i:end]

		ref, err := w.createNBTLeafPage(pageEntries)
		if err != nil {
			return nil, err
		}
		leafRefs = append(leafRefs, *ref)
		leafMinKeys = append(leafMinKeys, pageEntries[0].NID)
	}

	// Intermediate keys are the minimum key contained by each child subtree.
	return w.buildNBTNonleafLevels(leafRefs, leafMinKeys, 1)
}

// buildNBTNonleafLevels builds non-leaf levels of the NBT.
//
//nolint:dupl // NBT and BBT non-leaf building have same structure but different page types
func (w *BTWriter) buildNBTNonleafLevels(childRefs []disk.BlockReference, minKeys []uint64, level int) (*disk.BlockReference, error) {
	if len(childRefs) == 1 {
		return &childRefs[0], nil
	}

	format := w.db.Format()
	maxEntriesPerPage := disk.MaxBTPageEntriesNonleafUnicode
	if format == disk.FormatANSI {
		maxEntriesPerPage = disk.MaxBTPageEntriesNonleafANSI
	}

	// Create non-leaf entries
	var parentRefs []disk.BlockReference
	var parentMinKeys []uint64

	for i := 0; i < len(childRefs); i += maxEntriesPerPage {
		end := i + maxEntriesPerPage
		if end > len(childRefs) {
			end = len(childRefs)
		}

		// Build entries for this non-leaf page
		entries := make([]disk.BTNonleafEntry, end-i)
		for j := i; j < end; j++ {
			entries[j-i] = disk.BTNonleafEntry{
				Key: minKeys[j],
				Ref: childRefs[j],
			}
		}

		// Allocate and write page
		offset, err := w.amap.Allocate(disk.PageSize)
		if err != nil {
			return nil, err
		}

		bid := offset
		page := disk.NewBTNonleafPage(entries, level, disk.PageTypeNBT, bid)

		pageData, err := disk.BuildBTPage(page, format)
		if err != nil {
			return nil, err
		}

		if err := w.db.WritePage(offset, pageData); err != nil {
			return nil, err
		}

		parentRefs = append(parentRefs, disk.BlockReference{BID: bid, IB: offset})
		parentMinKeys = append(parentMinKeys, minKeys[i])
	}

	return w.buildNBTNonleafLevels(parentRefs, parentMinKeys, level+1)
}

// applyBBTChanges applies pending BBT changes.
func (w *BTWriter) applyBBTChanges() (*disk.BlockReference, error) {
	if len(w.bbtInserts) == 0 && len(w.bbtDeletes) == 0 {
		return &w.db.header.Root.BRefBBT, nil
	}

	// Load current BBT root
	if err := w.db.loadBBTRoot(); err != nil {
		return nil, err
	}

	// Collect all current entries
	entries, err := w.collectBBTEntries(w.db.bbtRoot)
	if err != nil {
		return nil, err
	}

	// Apply deletes
	filteredEntries := make([]disk.BBTLeafEntry, 0, len(entries))
	for _, entry := range entries {
		bid := util.BlockID(entry.BRef.BID) &^ util.BlockID(0x2)
		if !w.bbtDeletes[bid] {
			filteredEntries = append(filteredEntries, entry)
		}
	}

	// Apply inserts
	for _, info := range w.bbtInserts {
		entry := disk.BBTLeafEntry{
			BRef: disk.BlockReference{
				BID: uint64(info.BID),
				IB:  info.Location,
			},
			Size:     info.Size,
			RefCount: info.RefCount,
		}
		filteredEntries = append(filteredEntries, entry)
	}

	// Sort by BID
	sort.Slice(filteredEntries, func(i, j int) bool {
		left := filteredEntries[i].BRef.BID &^ uint64(0x2)
		right := filteredEntries[j].BRef.BID &^ uint64(0x2)
		return left < right
	})

	// Rebuild tree structure
	return w.rebuildBBT(filteredEntries)
}

// collectBBTEntries recursively collects all leaf entries from the BBT.
func (w *BTWriter) collectBBTEntries(page *disk.BTPage) ([]disk.BBTLeafEntry, error) {
	if page.IsLeaf() {
		return page.BBTEntries, nil
	}

	var entries []disk.BBTLeafEntry
	for _, child := range page.NonleafEntries {
		childData, err := w.db.readPage(child.Ref.IB)
		if err != nil {
			return nil, err
		}
		childPage, err := disk.ParseBTPage(childData, w.db.Format(), disk.PageTypeBBT)
		if err != nil {
			return nil, err
		}
		childEntries, err := w.collectBBTEntries(childPage)
		if err != nil {
			return nil, err
		}
		entries = append(entries, childEntries...)
	}
	return entries, nil
}

// rebuildBBT rebuilds the BBT from a sorted list of entries.
func (w *BTWriter) rebuildBBT(entries []disk.BBTLeafEntry) (*disk.BlockReference, error) {
	if len(entries) == 0 {
		return w.createEmptyBBT()
	}

	format := w.db.Format()
	maxEntriesPerPage := disk.MaxBTPageEntriesBBTUnicode
	if format == disk.FormatANSI {
		maxEntriesPerPage = disk.MaxBTPageEntriesBBTANSI
	}

	if len(entries) <= maxEntriesPerPage {
		return w.createBBTLeafPage(entries)
	}

	return w.buildBBTTree(entries, maxEntriesPerPage)
}

// createEmptyBBT creates an empty BBT.
func (w *BTWriter) createEmptyBBT() (*disk.BlockReference, error) {
	return w.createBBTLeafPage([]disk.BBTLeafEntry{})
}

// createBBTLeafPage creates a single BBT leaf page.
func (w *BTWriter) createBBTLeafPage(entries []disk.BBTLeafEntry) (*disk.BlockReference, error) {
	offset, err := w.amap.Allocate(disk.PageSize)
	if err != nil {
		return nil, err
	}

	bid := offset
	page := disk.NewBBTLeafPage(entries, bid, w.db.Format())

	pageData, err := disk.BuildBTPage(page, w.db.Format())
	if err != nil {
		return nil, err
	}

	if err := w.db.WritePage(offset, pageData); err != nil {
		return nil, err
	}

	w.newBBTRoot = page

	return &disk.BlockReference{
		BID: bid,
		IB:  offset,
	}, nil
}

// buildBBTTree builds a multi-level BBT from entries.
func (w *BTWriter) buildBBTTree(entries []disk.BBTLeafEntry, maxPerPage int) (*disk.BlockReference, error) {
	var leafRefs []disk.BlockReference
	var leafMinKeys []uint64

	for i := 0; i < len(entries); i += maxPerPage {
		end := i + maxPerPage
		if end > len(entries) {
			end = len(entries)
		}
		pageEntries := entries[i:end]

		ref, err := w.createBBTLeafPage(pageEntries)
		if err != nil {
			return nil, err
		}
		leafRefs = append(leafRefs, *ref)
		leafMinKeys = append(leafMinKeys, pageEntries[0].BRef.BID&^uint64(0x2))
	}

	return w.buildBBTNonleafLevels(leafRefs, leafMinKeys, 1)
}

// buildBBTNonleafLevels builds non-leaf levels of the BBT.
//
//nolint:dupl // NBT and BBT non-leaf building have same structure but different page types
func (w *BTWriter) buildBBTNonleafLevels(childRefs []disk.BlockReference, minKeys []uint64, level int) (*disk.BlockReference, error) {
	if len(childRefs) == 1 {
		return &childRefs[0], nil
	}

	format := w.db.Format()
	maxEntriesPerPage := disk.MaxBTPageEntriesNonleafUnicode
	if format == disk.FormatANSI {
		maxEntriesPerPage = disk.MaxBTPageEntriesNonleafANSI
	}

	var parentRefs []disk.BlockReference
	var parentMinKeys []uint64

	for i := 0; i < len(childRefs); i += maxEntriesPerPage {
		end := i + maxEntriesPerPage
		if end > len(childRefs) {
			end = len(childRefs)
		}

		entries := make([]disk.BTNonleafEntry, end-i)
		for j := i; j < end; j++ {
			entries[j-i] = disk.BTNonleafEntry{
				Key: minKeys[j],
				Ref: childRefs[j],
			}
		}

		offset, err := w.amap.Allocate(disk.PageSize)
		if err != nil {
			return nil, err
		}

		bid := offset
		page := disk.NewBTNonleafPage(entries, level, disk.PageTypeBBT, bid)

		pageData, err := disk.BuildBTPage(page, format)
		if err != nil {
			return nil, err
		}

		if err := w.db.WritePage(offset, pageData); err != nil {
			return nil, err
		}

		parentRefs = append(parentRefs, disk.BlockReference{BID: bid, IB: offset})
		parentMinKeys = append(parentMinKeys, minKeys[i])
	}

	return w.buildBBTNonleafLevels(parentRefs, parentMinKeys, level+1)
}

// HasChanges returns true if there are pending changes.
func (w *BTWriter) HasChanges() bool {
	return len(w.nbtInserts) > 0 || len(w.nbtDeletes) > 0 ||
		len(w.bbtInserts) > 0 || len(w.bbtDeletes) > 0
}

// Reset clears all pending changes.
func (w *BTWriter) Reset() {
	w.nbtInserts = make(map[util.NodeID]*NodeInfo)
	w.nbtDeletes = make(map[util.NodeID]bool)
	w.bbtInserts = make(map[util.BlockID]*BlockInfo)
	w.bbtDeletes = make(map[util.BlockID]bool)
	w.modifiedNBTPages = nil
	w.modifiedBBTPages = nil
	w.newNBTRoot = nil
	w.newBBTRoot = nil
}
