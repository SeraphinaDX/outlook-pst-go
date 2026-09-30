package outlookpst

import (
	"fmt"
	"iter"
	"sync"

	"github.com/grokify/outlook-pst-go/pkg/ltp"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// Folder represents a folder in a PST file.
type Folder struct {
	pst  *PST
	node *ndb.Node
	nid  util.NodeID
	bag  *ltp.PropertyBag

	// Hierarchy table (lazy loaded)
	hierarchyOnce  sync.Once
	hierarchyTable *ltp.Table
	hierarchyErr   error

	// Contents table (lazy loaded)
	contentsOnce  sync.Once
	contentsTable *ltp.Table
	contentsErr   error
}

// newFolder creates a new Folder from a node.
func newFolder(pst *PST, node *ndb.Node) (*Folder, error) {
	bag, err := ltp.NewPropertyBag(node)
	if err != nil {
		return nil, fmt.Errorf("failed to create property bag: %w", err)
	}

	return &Folder{
		pst:  pst,
		node: node,
		nid:  node.ID(),
		bag:  bag,
	}, nil
}

// ID returns the folder's node ID.
func (f *Folder) ID() util.NodeID {
	if f.node != nil {
		return f.node.ID()
	}
	return f.nid
}

// Name returns the display name of the folder.
func (f *Folder) Name() (string, error) {
	return f.bag.GetString(ltp.PidTagDisplayName)
}

// ContentCount returns the number of messages in the folder.
func (f *Folder) ContentCount() (int32, error) {
	return f.bag.GetInt32(ltp.PidTagContentCount)
}

// UnreadCount returns the number of unread messages.
func (f *Folder) UnreadCount() (int32, error) {
	return f.bag.GetInt32(ltp.PidTagContentUnreadCount)
}

// HasSubfolders returns true if the folder has subfolders.
func (f *Folder) HasSubfolders() (bool, error) {
	return f.bag.GetBool(ltp.PidTagSubfolders)
}

// ContainerClass returns the folder's container class (e.g., "IPF.Note").
func (f *Folder) ContainerClass() (string, error) {
	return f.bag.GetString(ltp.PidTagContainerClass)
}

// PropertyBag returns the folder's property bag for advanced access.
func (f *Folder) PropertyBag() *ltp.PropertyBag {
	return f.bag
}

// loadHierarchyTable loads the hierarchy table for subfolders.
func (f *Folder) loadHierarchyTable() error {
	f.hierarchyOnce.Do(func() {
		// Hierarchy table NID is folder NID with type changed to hierarchy
		folderNID := f.ID()
		hierarchyNID := util.MakeNID(util.NIDTypeHierarchyTable, folderNID.Index())

		// The hierarchy table is a separate NBT entry (not a subnode)
		hierarchyNode, err := f.pst.db.GetNode(hierarchyNID)
		if err != nil {
			f.hierarchyErr = fmt.Errorf("hierarchy table not found: %w", err)
			return
		}

		f.hierarchyTable, err = ltp.NewTable(hierarchyNode)
		if err != nil {
			f.hierarchyErr = fmt.Errorf("failed to create hierarchy table: %w", err)
			return
		}
	})
	return f.hierarchyErr
}

// loadContentsTable loads the contents table for messages.
func (f *Folder) loadContentsTable() error {
	f.contentsOnce.Do(func() {
		// Contents table NID is folder NID with type changed to contents
		folderNID := f.ID()
		contentsNID := util.MakeNID(util.NIDTypeContentsTable, folderNID.Index())

		// The contents table is a separate NBT entry (not a subnode)
		contentsNode, err := f.pst.db.GetNode(contentsNID)
		if err != nil {
			f.contentsErr = fmt.Errorf("contents table not found: %w", err)
			return
		}

		f.contentsTable, err = ltp.NewTable(contentsNode)
		if err != nil {
			f.contentsErr = fmt.Errorf("failed to create contents table: %w", err)
			return
		}
	})
	return f.contentsErr
}

// Subfolders returns an iterator over subfolders.
func (f *Folder) Subfolders() iter.Seq2[*Folder, error] {
	return func(yield func(*Folder, error) bool) {
		if err := f.loadHierarchyTable(); err != nil {
			yield(nil, err)
			return
		}

		for row, err := range f.hierarchyTable.Rows() {
			if err != nil {
				yield(nil, err)
				return
			}

			// The row ID (dwRowID) IS the subfolder NID.
			// It comes from the BTH key, not the row data.
			// See [MS-PST] Section 2.3.4.3 - The dwRowID for hierarchy tables
			// is the NID of the child folder.
			subfolderNID := util.NodeID(row.RowID())

			// Get the subfolder node
			subfolderNode, err := f.pst.db.GetNode(subfolderNID)
			if err != nil {
				yield(nil, fmt.Errorf("failed to get subfolder node: %w", err))
				return
			}

			subfolder, err := newFolder(f.pst, subfolderNode)
			if err != nil {
				yield(nil, err)
				return
			}

			if !yield(subfolder, nil) {
				return
			}
		}
	}
}

// SubfolderCount returns the number of subfolders.
func (f *Folder) SubfolderCount() (int, error) {
	if err := f.loadHierarchyTable(); err != nil {
		return 0, err
	}
	return f.hierarchyTable.RowCount(), nil
}

// FindSubfolder finds a subfolder by name.
func (f *Folder) FindSubfolder(name string) (*Folder, error) {
	for subfolder, err := range f.Subfolders() {
		if err != nil {
			return nil, err
		}

		subName, err := subfolder.Name()
		if err != nil {
			continue
		}

		if subName == name {
			return subfolder, nil
		}
	}

	return nil, &FolderNotFoundError{Name: name}
}

// Messages returns an iterator over messages in the folder.
func (f *Folder) Messages() iter.Seq2[*Message, error] {
	return func(yield func(*Message, error) bool) {
		if err := f.loadContentsTable(); err != nil {
			yield(nil, err)
			return
		}

		for row, err := range f.contentsTable.Rows() {
			if err != nil {
				yield(nil, err)
				return
			}

			// The row ID (dwRowID) IS the message NID.
			// It comes from the BTH key, not the row data.
			// See [MS-PST] Section 2.3.4.3 - The dwRowID for contents tables
			// is the NID of the message.
			messageNID := util.NodeID(row.RowID())

			// Get the message node
			messageNode, err := f.pst.db.GetNode(messageNID)
			if err != nil {
				yield(nil, fmt.Errorf("failed to get message node: %w", err))
				return
			}

			message, err := newMessage(f.pst, messageNode)
			if err != nil {
				yield(nil, err)
				return
			}

			if !yield(message, nil) {
				return
			}
		}
	}
}

// MessageCount returns the number of messages.
func (f *Folder) MessageCount() (int, error) {
	if err := f.loadContentsTable(); err != nil {
		return 0, err
	}
	return f.contentsTable.RowCount(), nil
}

// HierarchyTable returns the hierarchy table for advanced access.
func (f *Folder) HierarchyTable() (*ltp.Table, error) {
	if err := f.loadHierarchyTable(); err != nil {
		return nil, err
	}
	return f.hierarchyTable, nil
}

// ContentsTable returns the contents table for advanced access.
func (f *Folder) ContentsTable() (*ltp.Table, error) {
	if err := f.loadContentsTable(); err != nil {
		return nil, err
	}
	return f.contentsTable, nil
}

// IsSearchFolder returns true if this is a search folder.
// Search folders are virtual folders whose contents are determined by search criteria.
// See [MS-PST] Section 2.4.8.
func (f *Folder) IsSearchFolder() bool {
	return f.ID().Type() == util.NIDTypeSearchFolder
}

// AsSearchFolder returns this folder as a SearchFolder if it is one.
// Returns nil if this is not a search folder.
func (f *Folder) AsSearchFolder() *SearchFolder {
	if !f.IsSearchFolder() {
		return nil
	}
	return newSearchFolder(f)
}
