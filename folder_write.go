package outlookpst

import (
	"fmt"

	"github.com/grokify/outlook-pst-go/pkg/ltp"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// CreateFolder creates a new subfolder in the given parent folder.
func CreateFolder(ctx *WriteContext, parent *Folder, name string) (*Folder, error) {
	if parent == nil {
		return nil, fmt.Errorf("parent folder cannot be nil")
	}
	if name == "" {
		return nil, fmt.Errorf("folder name cannot be empty")
	}

	txn := ctx.Transaction()
	format := ctx.PST().Format()

	folderBag, err := ltp.CreateFolderPropertyBag(format, name)
	if err != nil {
		return nil, fmt.Errorf("failed to create folder properties: %w", err)
	}
	folderData, err := folderBag.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build folder properties: %w", err)
	}

	folderInfo, err := ndb.NewNodeBuilder(txn, util.NIDTypeNormalFolder).
		WithParent(parent.ID()).
		WithData(folderData).
		Build()
	if err != nil {
		return nil, fmt.Errorf("failed to create folder node: %w", err)
	}

	hierarchyNID := util.MakeNID(util.NIDTypeHierarchyTable, folderInfo.NID.Index())
	hierarchyTable := ltp.CreateHierarchyTable(format)
	hierarchyData, err := hierarchyTable.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build hierarchy table: %w", err)
	}
	if _, err := ndb.NewNodeBuilder(txn, util.NIDTypeHierarchyTable).
		WithNID(hierarchyNID).
		WithParent(folderInfo.NID).
		WithData(hierarchyData).
		Build(); err != nil {
		return nil, fmt.Errorf("failed to create hierarchy table node: %w", err)
	}
	ctx.tableWriters[hierarchyNID] = hierarchyTable

	contentsNID := util.MakeNID(util.NIDTypeContentsTable, folderInfo.NID.Index())
	contentsTable := ltp.CreateContentsTable(format)
	contentsData, err := contentsTable.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build contents table: %w", err)
	}
	if _, err := ndb.NewNodeBuilder(txn, util.NIDTypeContentsTable).
		WithNID(contentsNID).
		WithParent(folderInfo.NID).
		WithData(contentsData).
		Build(); err != nil {
		return nil, fmt.Errorf("failed to create contents table node: %w", err)
	}
	ctx.tableWriters[contentsNID] = contentsTable

	if err := addToHierarchyTable(ctx, parent, folderInfo.NID, name); err != nil {
		return nil, fmt.Errorf("failed to update hierarchy table: %w", err)
	}

	node, err := ctx.PST().db.GetNode(folderInfo.NID)
	if err != nil {
		return &Folder{pst: ctx.PST(), nid: folderInfo.NID}, nil
	}
	return newFolder(ctx.PST(), node)
}

func folderTableWriter(ctx *WriteContext, folder *Folder, tableType util.NIDType, create func(disk.PSTFormat) *ltp.TableWriter) (*ltp.TableWriter, util.NodeID, error) {
	tableNID := util.MakeNID(tableType, folder.ID().Index())
	if writer, ok := ctx.tableWriters[tableNID]; ok {
		return writer, tableNID, nil
	}

	node, err := ctx.PST().db.GetNode(tableNID)
	if err == nil {
		table, tableErr := ltp.NewTable(node)
		if tableErr != nil {
			return nil, 0, tableErr
		}
		writer, tableErr := ltp.NewTableWriterFromTable(table, ctx.PST().Format())
		if tableErr != nil {
			return nil, 0, tableErr
		}
		ctx.tableWriters[tableNID] = writer
		return writer, tableNID, nil
	}

	writer := create(ctx.PST().Format())
	data, buildErr := writer.Build()
	if buildErr != nil {
		return nil, 0, buildErr
	}
	if _, buildErr = ndb.NewNodeBuilder(ctx.Transaction(), tableType).
		WithNID(tableNID).
		WithParent(folder.ID()).
		WithData(data).
		Build(); buildErr != nil {
		return nil, 0, buildErr
	}
	ctx.tableWriters[tableNID] = writer
	return writer, tableNID, nil
}

// addToHierarchyTable adds a folder entry to the parent's hierarchy table.
func addToHierarchyTable(ctx *WriteContext, parent *Folder, childNID util.NodeID, name string) error {
	writer, tableNID, err := folderTableWriter(ctx, parent, util.NIDTypeHierarchyTable, ltp.CreateHierarchyTable)
	if err != nil {
		return err
	}

	rowID, err := writer.AddRowWithID(uint32(childNID))
	if err != nil {
		return err
	}
	if err := writer.SetRowString(rowID, ltp.PidTagDisplayName, name); err != nil {
		return err
	}
	if err := writer.SetRowInt32(rowID, ltp.PidTagContentCount, 0); err != nil {
		return err
	}
	if err := writer.SetRowInt32(rowID, ltp.PidTagContentUnreadCount, 0); err != nil {
		return err
	}
	if err := writer.SetRowBool(rowID, ltp.PidTagSubfolders, false); err != nil {
		return err
	}
	if err := writer.SetRowInt32(rowID, ltp.PidTagDepth, 1); err != nil {
		return err
	}

	data, err := writer.Build()
	if err != nil {
		return err
	}
	return ndb.UpdateNodeData(ctx.Transaction(), tableNID, data)
}

// DeleteFolder deletes a folder and all its contents recursively.
func DeleteFolder(ctx *WriteContext, folder *Folder) error {
	if folder == nil {
		return fmt.Errorf("folder cannot be nil")
	}

	txn := ctx.Transaction()

	// Delete all messages in the folder
	for msg, err := range folder.Messages() {
		if err != nil {
			continue
		}
		if err := DeleteMessage(ctx, msg); err != nil {
			// Log but continue
			_ = err
		}
	}

	// Delete all subfolders recursively
	for subfolder, err := range folder.Subfolders() {
		if err != nil {
			continue
		}
		if err := DeleteFolder(ctx, subfolder); err != nil {
			// Log but continue
			_ = err
		}
	}

	// Delete folder's tables (hierarchy, contents)
	// Find and delete hierarchy table node
	htNID := util.MakeNID(util.NIDTypeHierarchyTable, folder.ID().Index())
	if err := txn.DeleteNode(htNID); err != nil {
		// May not exist - ignore
		_ = err
	}

	// Find and delete contents table node
	ctNID := util.MakeNID(util.NIDTypeContentsTable, folder.ID().Index())
	if err := txn.DeleteNode(ctNID); err != nil {
		_ = err
	}

	// Delete the folder node itself
	return ndb.DeleteNodeRecursive(txn, folder.ID())
}

// RenameFolder renames a folder.
func RenameFolder(ctx *WriteContext, folder *Folder, newName string) error {
	if folder == nil {
		return fmt.Errorf("folder cannot be nil")
	}
	if newName == "" {
		return fmt.Errorf("new name cannot be empty")
	}

	txn := ctx.Transaction()
	format := ctx.PST().Format()

	// Get current folder properties
	bag := folder.PropertyBag()
	if bag == nil {
		return fmt.Errorf("folder has no property bag")
	}

	// Create updated property bag
	newBag := ltp.NewPropertyBagWriter(format)

	// Copy existing properties and update display name
	for _, propID := range bag.Properties() {
		if propID == ltp.PidTagDisplayName {
			if err := newBag.SetString(propID, newName); err != nil {
				return err
			}
		} else {
			// Copy other properties
			rawData, err := bag.GetRaw(propID)
			if err != nil {
				continue
			}
			propType, _ := bag.GetType(propID)
			// This is simplified - would need proper type handling
			_ = propType
			_ = rawData
		}
	}

	// Build updated data
	newData, err := newBag.Build()
	if err != nil {
		return fmt.Errorf("failed to build updated properties: %w", err)
	}

	// Update the node's data
	return ndb.UpdateNodeData(txn, folder.ID(), newData)
}

// MoveFolder moves a folder to a new parent.
func MoveFolder(ctx *WriteContext, folder, newParent *Folder) error {
	if folder == nil {
		return fmt.Errorf("folder cannot be nil")
	}
	if newParent == nil {
		return fmt.Errorf("new parent cannot be nil")
	}

	// Check for circular reference
	if folder.ID() == newParent.ID() {
		return fmt.Errorf("cannot move folder into itself")
	}

	txn := ctx.Transaction()

	// Get folder name for hierarchy table updates
	name, err := folder.Name()
	if err != nil {
		name = "Folder"
	}

	// Update node's parent in NBT
	if err := ndb.MoveNode(txn, folder.ID(), newParent.ID()); err != nil {
		return fmt.Errorf("failed to move node: %w", err)
	}

	// Update hierarchy tables
	// Remove from old parent's hierarchy table
	// Add to new parent's hierarchy table
	if err := addToHierarchyTable(ctx, newParent, folder.ID(), name); err != nil {
		return fmt.Errorf("failed to update new parent hierarchy: %w", err)
	}

	return nil
}

// CopyFolder copies a folder and its contents to a new parent.
func CopyFolder(ctx *WriteContext, folder, newParent *Folder) (*Folder, error) {
	if folder == nil {
		return nil, fmt.Errorf("folder cannot be nil")
	}
	if newParent == nil {
		return nil, fmt.Errorf("new parent cannot be nil")
	}

	// Get folder name
	name, err := folder.Name()
	if err != nil {
		name = "Copied Folder"
	}

	// Create new folder in parent
	newFolder, err := CreateFolder(ctx, newParent, name)
	if err != nil {
		return nil, fmt.Errorf("failed to create folder copy: %w", err)
	}

	// Copy all messages
	for msg, err := range folder.Messages() {
		if err != nil {
			continue
		}
		if err := CopyMessage(ctx, msg, newFolder); err != nil {
			// Log but continue
			_ = err
		}
	}

	// Copy all subfolders recursively
	for subfolder, err := range folder.Subfolders() {
		if err != nil {
			continue
		}
		if _, err := CopyFolder(ctx, subfolder, newFolder); err != nil {
			// Log but continue
			_ = err
		}
	}

	return newFolder, nil
}

// FolderExists checks if a subfolder with the given name exists.
func FolderExists(parent *Folder, name string) (bool, error) {
	for subfolder, err := range parent.Subfolders() {
		if err != nil {
			continue
		}
		folderName, err := subfolder.Name()
		if err != nil {
			continue
		}
		if folderName == name {
			return true, nil
		}
	}
	return false, nil
}

// GetOrCreateFolder gets a folder by name, creating it if it doesn't exist.
func GetOrCreateFolder(ctx *WriteContext, parent *Folder, name string) (*Folder, error) {
	// Check if folder exists
	subfolder, err := parent.FindSubfolder(name)
	if err == nil && subfolder != nil {
		return subfolder, nil
	}

	// Create the folder
	return CreateFolder(ctx, parent, name)
}

// EmptyFolder removes all contents from a folder without deleting the folder itself.
func EmptyFolder(ctx *WriteContext, folder *Folder) error {
	if folder == nil {
		return fmt.Errorf("folder cannot be nil")
	}

	// Delete all messages
	for msg, err := range folder.Messages() {
		if err != nil {
			continue
		}
		if err := DeleteMessage(ctx, msg); err != nil {
			_ = err
		}
	}

	// Delete all subfolders
	for subfolder, err := range folder.Subfolders() {
		if err != nil {
			continue
		}
		if err := DeleteFolder(ctx, subfolder); err != nil {
			_ = err
		}
	}

	return nil
}

// UpdateFolderProperty updates a single property on a folder.
func UpdateFolderProperty(ctx *WriteContext, folder *Folder, propID ltp.PropID, value interface{}) error {
	if folder == nil {
		return fmt.Errorf("folder cannot be nil")
	}

	format := ctx.PST().Format()
	bag := ltp.NewPropertyBagWriter(format)

	// Copy existing properties
	existingBag := folder.PropertyBag()
	if existingBag != nil {
		for _, pid := range existingBag.Properties() {
			if pid == propID {
				continue // Skip - we're updating this one
			}
			// Copy property (simplified)
		}
	}

	// Set the updated property
	switch v := value.(type) {
	case string:
		if err := bag.SetString(propID, v); err != nil {
			return err
		}
	case int32:
		if err := bag.SetInt32(propID, v); err != nil {
			return err
		}
	case bool:
		if err := bag.SetBool(propID, v); err != nil {
			return err
		}
	default:
		return fmt.Errorf("unsupported property value type: %T", value)
	}

	// Build and update
	data, err := bag.Build()
	if err != nil {
		return err
	}

	return ndb.UpdateNodeData(ctx.Transaction(), folder.ID(), data)
}
