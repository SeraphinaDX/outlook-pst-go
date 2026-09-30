package outlookpst

import (
	"fmt"
	"time"

	"github.com/grokify/outlook-pst-go/pkg/ltp"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// AttachMethod indicates how an attachment is stored.
type AttachMethod int32

const (
	AttachMethodNone         AttachMethod = 0
	AttachMethodByValue      AttachMethod = 1
	AttachMethodByReference  AttachMethod = 2
	AttachMethodByRefResolve AttachMethod = 3
	AttachMethodByRefOnly    AttachMethod = 4
	AttachMethodEmbedded     AttachMethod = 5
	AttachMethodOLE          AttachMethod = 6
)

// MessageBuilder provides a fluent interface for building messages.
type MessageBuilder struct {
	ctx         *WriteContext
	folder      *Folder
	subject     string
	body        string
	htmlBody    string
	rtfBody     []byte
	sentTime    time.Time
	fromName    string
	fromEmail   string
	recipients  []recipientInfo
	attachments []attachmentInfo
	properties  map[ltp.PropID]interface{}
}

// recipientInfo holds recipient information.
type recipientInfo struct {
	recipientType RecipientType
	displayName   string
	emailAddress  string
	addressType   string
}

// attachmentInfo holds attachment information.
type attachmentInfo struct {
	filename string
	mimeType string
	data     []byte
	method   AttachMethod
}

// NewMessageBuilder creates a new message builder.
func NewMessageBuilder(ctx *WriteContext, folder *Folder) *MessageBuilder {
	return &MessageBuilder{
		ctx:        ctx,
		folder:     folder,
		sentTime:   time.Now(),
		properties: make(map[ltp.PropID]interface{}),
	}
}

// SetSubject sets the message subject.
func (b *MessageBuilder) SetSubject(subject string) *MessageBuilder {
	b.subject = subject
	return b
}

// SetBody sets the plain text body.
func (b *MessageBuilder) SetBody(body string) *MessageBuilder {
	b.body = body
	return b
}

// SetHTMLBody sets the HTML body.
func (b *MessageBuilder) SetHTMLBody(html string) *MessageBuilder {
	b.htmlBody = html
	return b
}

// SetRTFBody sets the RTF body.
func (b *MessageBuilder) SetRTFBody(rtf []byte) *MessageBuilder {
	b.rtfBody = rtf
	return b
}

// SetSentTime sets the sent time.
func (b *MessageBuilder) SetSentTime(t time.Time) *MessageBuilder {
	b.sentTime = t
	return b
}

// SetFrom sets the sender information.
func (b *MessageBuilder) SetFrom(name, email string) *MessageBuilder {
	b.fromName = name
	b.fromEmail = email
	return b
}

// AddRecipient adds a recipient to the message.
func (b *MessageBuilder) AddRecipient(name, email string, rtype RecipientType) *MessageBuilder {
	b.recipients = append(b.recipients, recipientInfo{
		recipientType: rtype,
		displayName:   name,
		emailAddress:  email,
		addressType:   "SMTP",
	})
	return b
}

// AddTo adds a TO recipient.
func (b *MessageBuilder) AddTo(name, email string) *MessageBuilder {
	return b.AddRecipient(name, email, RecipientTo)
}

// AddCC adds a CC recipient.
func (b *MessageBuilder) AddCC(name, email string) *MessageBuilder {
	return b.AddRecipient(name, email, RecipientCc)
}

// AddBCC adds a BCC recipient.
func (b *MessageBuilder) AddBCC(name, email string) *MessageBuilder {
	return b.AddRecipient(name, email, RecipientBcc)
}

// AddAttachment adds a file attachment.
func (b *MessageBuilder) AddAttachment(filename string, data []byte) *MessageBuilder {
	return b.AddAttachmentWithMime(filename, data, "application/octet-stream")
}

// AddAttachmentWithMime adds an attachment with a specific MIME type.
func (b *MessageBuilder) AddAttachmentWithMime(filename string, data []byte, mimeType string) *MessageBuilder {
	b.attachments = append(b.attachments, attachmentInfo{
		filename: filename,
		mimeType: mimeType,
		data:     data,
		method:   AttachMethodByValue,
	})
	return b
}

// SetProperty sets a custom property.
func (b *MessageBuilder) SetProperty(propID ltp.PropID, value interface{}) *MessageBuilder {
	b.properties[propID] = value
	return b
}

// Build creates the message in the PST file.
func (b *MessageBuilder) Build() (*Message, error) {
	if b.folder == nil {
		return nil, fmt.Errorf("folder cannot be nil")
	}

	txn := b.ctx.Transaction()
	format := b.ctx.PST().Format()

	// Reserve the message NID before building subnodes. Recipient and
	// attachment table NIDs share the message's nidIndex.
	msgNID, err := txn.AllocateNodeID(util.NIDTypeNormalMessage)
	if err != nil {
		return nil, err
	}

	msgBag := ltp.NewPropertyBagWriter(format)
	if err := msgBag.SetString(ltp.PidTagMessageClass, "IPM.Note"); err != nil {
		return nil, err
	}
	if b.subject != "" {
		if err := msgBag.SetString(ltp.PidTagSubject, b.subject); err != nil {
			return nil, err
		}
		if err := msgBag.SetString(ltp.PidTagNormalizedSubject, b.subject); err != nil {
			return nil, err
		}
	}
	if b.body != "" {
		if err := msgBag.SetString(ltp.PidTagBody, b.body); err != nil {
			return nil, err
		}
	}
	if b.htmlBody != "" {
		if err := msgBag.SetString(ltp.PidTagHtmlBody, b.htmlBody); err != nil {
			return nil, err
		}
	}
	if len(b.rtfBody) > 0 {
		if err := msgBag.SetBinary(ltp.PidTagRtfCompressed, b.rtfBody); err != nil {
			return nil, err
		}
	}

	if err := msgBag.SetTime(ltp.PidTagClientSubmitTime, b.sentTime); err != nil {
		return nil, err
	}
	if err := msgBag.SetTime(ltp.PidTagMessageDeliveryTime, b.sentTime); err != nil {
		return nil, err
	}
	now := time.Now()
	if err := msgBag.SetTime(ltp.PidTagCreationTime, now); err != nil {
		return nil, err
	}
	if err := msgBag.SetTime(ltp.PidTagLastModificationTime, now); err != nil {
		return nil, err
	}

	if b.fromName != "" {
		if err := msgBag.SetString(ltp.PidTagSenderName, b.fromName); err != nil {
			return nil, err
		}
	}
	if b.fromEmail != "" {
		if err := msgBag.SetString(ltp.PidTagSenderEmailAddress, b.fromEmail); err != nil {
			return nil, err
		}
		if err := msgBag.SetString(ltp.PidTagSenderAddressType, "SMTP"); err != nil {
			return nil, err
		}
	}

	flags := int32(1)
	if len(b.attachments) > 0 {
		flags |= 0x10
	}
	if err := msgBag.SetInt32(ltp.PidTagMessageFlags, flags); err != nil {
		return nil, err
	}
	if err := msgBag.SetBool(ltp.PidTagHasAttachments, len(b.attachments) > 0); err != nil {
		return nil, err
	}

	for propID, value := range b.properties {
		switch v := value.(type) {
		case string:
			err = msgBag.SetString(propID, v)
		case int32:
			err = msgBag.SetInt32(propID, v)
		case int64:
			err = msgBag.SetInt64(propID, v)
		case bool:
			err = msgBag.SetBool(propID, v)
		case time.Time:
			err = msgBag.SetTime(propID, v)
		case []byte:
			err = msgBag.SetBinary(propID, v)
		}
		if err != nil {
			return nil, err
		}
	}

	subnodeBuilder := ndb.NewSubnodeBuilder(txn)
	if err := spillLargeProperties(msgBag, subnodeBuilder); err != nil {
		return nil, fmt.Errorf("failed to externalize message properties: %w", err)
	}

	msgData, err := msgBag.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build message properties: %w", err)
	}

	if len(b.recipients) > 0 {
		recipientTable := ltp.CreateRecipientTable(format)
		for i, r := range b.recipients {
			rowID := recipientTable.AddRow()
			if err := recipientTable.SetRowInt32(rowID, ltp.PidTagRecipientType, int32(r.recipientType)); err != nil {
				return nil, err
			}
			if err := recipientTable.SetRowString(rowID, ltp.PidTagDisplayName, r.displayName); err != nil {
				return nil, err
			}
			if err := recipientTable.SetRowString(rowID, ltp.PidTagEmailAddress, r.emailAddress); err != nil {
				return nil, err
			}
			if err := recipientTable.SetRowString(rowID, ltp.PidTagAddressType, r.addressType); err != nil {
				return nil, err
			}
			_ = i
		}
		recipientData, err := recipientTable.Build()
		if err != nil {
			return nil, fmt.Errorf("failed to build recipient table: %w", err)
		}
		recipientNID := util.MakeNID(util.NIDTypeRecipientTable, msgNID.Index())
		if err := subnodeBuilder.AddSubnode(recipientNID, recipientData); err != nil {
			return nil, err
		}
	}

	if len(b.attachments) > 0 {
		attachmentTable := ltp.CreateAttachmentTable(format)
		for i, a := range b.attachments {
			attachNID := util.MakeNID(util.NIDTypeAttachment, uint32(i+1))
			rowID, err := attachmentTable.AddRowWithID(uint32(attachNID))
			if err != nil {
				return nil, err
			}
			if err := attachmentTable.SetRowInt32(rowID, ltp.PidTagAttachNumber, int32(i)); err != nil {
				return nil, err
			}
			if err := attachmentTable.SetRowInt32(rowID, ltp.PidTagAttachMethod, int32(a.method)); err != nil {
				return nil, err
			}
			if err := attachmentTable.SetRowString(rowID, ltp.PidTagAttachFilename, a.filename); err != nil {
				return nil, err
			}
			if err := attachmentTable.SetRowInt32(rowID, ltp.PidTagAttachSize, int32(len(a.data))); err != nil { //nolint:gosec
				return nil, err
			}
			if err := attachmentTable.SetRowString(rowID, ltp.PidTagAttachMimeTag, a.mimeType); err != nil {
				return nil, err
			}

			attachBag := ltp.NewPropertyBagWriter(format)
			_ = attachBag.SetInt32(ltp.PidTagAttachNumber, int32(i))
			_ = attachBag.SetInt32(ltp.PidTagAttachMethod, int32(a.method))
			_ = attachBag.SetString(ltp.PidTagAttachFilename, a.filename)
			_ = attachBag.SetString(ltp.PidTagAttachLongFilename, a.filename)
			_ = attachBag.SetBinary(ltp.PidTagAttachDataBinary, a.data)
			_ = attachBag.SetInt32(ltp.PidTagAttachSize, int32(len(a.data))) //nolint:gosec
			_ = attachBag.SetString(ltp.PidTagAttachMimeTag, a.mimeType)

			attachSubnodes := ndb.NewSubnodeBuilder(txn)
			if err := spillLargeProperties(attachBag, attachSubnodes); err != nil {
				return nil, fmt.Errorf("failed to externalize attachment %d properties: %w", i, err)
			}
			attachData, err := attachBag.Build()
			if err != nil {
				return nil, fmt.Errorf("failed to build attachment %d: %w", i, err)
			}
			attachSubBID, err := attachSubnodes.Build()
			if err != nil {
				return nil, fmt.Errorf("failed to build attachment %d subnodes: %w", i, err)
			}
			if err := subnodeBuilder.AddSubnodeWithSubnodes(attachNID, attachData, attachSubBID); err != nil {
				return nil, err
			}
		}

		attachmentData, err := attachmentTable.Build()
		if err != nil {
			return nil, fmt.Errorf("failed to build attachment table: %w", err)
		}
		attachmentTableNID := util.MakeNID(util.NIDTypeAttachmentTable, msgNID.Index())
		if err := subnodeBuilder.AddSubnode(attachmentTableNID, attachmentData); err != nil {
			return nil, err
		}
	}

	subBID, err := subnodeBuilder.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build subnodes: %w", err)
	}

	msgInfo, err := ndb.NewNodeBuilder(txn, util.NIDTypeNormalMessage).
		WithNID(msgNID).
		WithParent(b.folder.ID()).
		WithData(msgData).
		WithSubnodeBID(subBID).
		Build()
	if err != nil {
		return nil, fmt.Errorf("failed to create message node: %w", err)
	}

	if err := addToContentsTable(b.ctx, b.folder, msgInfo.NID, b.subject, b.sentTime, flags, len(b.attachments) > 0); err != nil {
		return nil, fmt.Errorf("failed to update contents table: %w", err)
	}

	node, err := b.ctx.PST().db.GetNode(msgInfo.NID)
	if err != nil {
		return &Message{pst: b.ctx.PST(), nid: msgInfo.NID}, nil
	}
	return newMessage(b.ctx.PST(), node)
}

func spillLargeProperties(bag *ltp.PropertyBagWriter, subnodes *ndb.SubnodeBuilder) error {
	for _, prop := range bag.ExternalProperties() {
		// Subnode NIDs are scoped to their parent node, so deriving the LTP
		// nidIndex from the property ID gives a stable unique reference.
		nid := util.MakeNID(util.NIDTypeLTP, uint32(prop.ID))
		if err := subnodes.AddSubnode(nid, prop.Data); err != nil {
			return fmt.Errorf("property 0x%04X: %w", uint16(prop.ID), err)
		}
		if err := bag.SetSubnodeReference(prop.ID, nid); err != nil {
			return err
		}
	}
	return nil
}

// addToContentsTable adds a message entry to the folder's contents table.
func addToContentsTable(ctx *WriteContext, folder *Folder, msgNID util.NodeID, subject string, sentTime time.Time, flags int32, hasAttachments bool) error {
	writer, tableNID, err := folderTableWriter(ctx, folder, util.NIDTypeContentsTable, ltp.CreateContentsTable)
	if err != nil {
		return err
	}

	rowID, err := writer.AddRowWithID(uint32(msgNID))
	if err != nil {
		return err
	}
	if subject != "" {
		if err := writer.SetRowString(rowID, ltp.PidTagSubject, subject); err != nil {
			return err
		}
	}
	if err := writer.SetRowString(rowID, ltp.PidTagMessageClass, "IPM.Note"); err != nil {
		return err
	}
	if err := writer.SetRowTime(rowID, ltp.PidTagClientSubmitTime, ltp.TimeToFileTime(sentTime)); err != nil {
		return err
	}
	if err := writer.SetRowInt32(rowID, ltp.PidTagMessageSize, 0); err != nil {
		return err
	}
	if err := writer.SetRowInt32(rowID, ltp.PidTagMessageFlags, flags); err != nil {
		return err
	}
	if err := writer.SetRowBool(rowID, ltp.PidTagHasAttachments, hasAttachments); err != nil {
		return err
	}

	data, err := writer.Build()
	if err != nil {
		return err
	}
	return ndb.UpdateNodeData(ctx.Transaction(), tableNID, data)
}

// DeleteMessage deletes a message from a folder.
func DeleteMessage(ctx *WriteContext, msg *Message) error {
	if msg == nil {
		return fmt.Errorf("message cannot be nil")
	}

	txn := ctx.Transaction()

	// Delete the message node and all subnodes
	return ndb.DeleteNodeRecursive(txn, msg.ID())
}

// CopyMessage copies a message to another folder.
func CopyMessage(ctx *WriteContext, msg *Message, destFolder *Folder) error {
	if msg == nil {
		return fmt.Errorf("message cannot be nil")
	}
	if destFolder == nil {
		return fmt.Errorf("destination folder cannot be nil")
	}

	// Get message properties
	subject, _ := msg.Subject()
	body, _ := msg.Body()
	sentTime, _ := msg.SubmitTime()
	htmlBody, _ := msg.HTMLBody()

	// Create copy in destination
	builder := ctx.CreateMessage(destFolder)
	builder.SetSubject(subject)
	builder.SetBody(body)
	builder.SetHTMLBody(htmlBody)
	builder.SetSentTime(sentTime)

	// Copy recipients
	for recip, err := range msg.Recipients() {
		if err != nil {
			continue
		}
		name, _ := recip.Name()
		email, _ := recip.Email()
		rtype, _ := recip.Type()
		builder.AddRecipient(name, email, rtype)
	}

	// Copy attachments
	for attach, err := range msg.Attachments() {
		if err != nil {
			continue
		}
		filename, _ := attach.Filename()
		data, _ := attach.Data()
		builder.AddAttachment(filename, data)
	}

	_, err := builder.Build()
	return err
}

// MoveMessage moves a message to another folder.
func MoveMessage(ctx *WriteContext, msg *Message, destFolder *Folder) error {
	// Copy then delete
	if err := CopyMessage(ctx, msg, destFolder); err != nil {
		return fmt.Errorf("copy failed: %w", err)
	}
	return DeleteMessage(ctx, msg)
}

// UpdateMessage updates properties on an existing message.
func UpdateMessage(ctx *WriteContext, msg *Message, updates map[ltp.PropID]interface{}) error {
	if msg == nil {
		return fmt.Errorf("message cannot be nil")
	}

	format := ctx.PST().Format()
	bag := ltp.NewPropertyBagWriter(format)

	// Copy existing properties
	existingBag := msg.PropertyBag()
	if existingBag != nil {
		for _, propID := range existingBag.Properties() {
			// Skip properties that will be updated
			if updates[propID] != nil {
				continue
			}
			// Copy property (simplified)
		}
	}

	// Apply updates
	for propID, value := range updates {
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
		case time.Time:
			if err := bag.SetTime(propID, v); err != nil {
				return err
			}
		case []byte:
			if err := bag.SetBinary(propID, v); err != nil {
				return err
			}
		}
	}

	// Update modification time
	if err := bag.SetTime(ltp.PidTagLastModificationTime, time.Now()); err != nil {
		return err
	}

	// Build and update
	data, err := bag.Build()
	if err != nil {
		return err
	}

	return ndb.UpdateNodeData(ctx.Transaction(), msg.ID(), data)
}

// MarkAsRead marks a message as read.
func MarkAsRead(ctx *WriteContext, msg *Message) error {
	return UpdateMessage(ctx, msg, map[ltp.PropID]interface{}{
		ltp.PidTagMessageFlags: int32(1), // MSGFLAG_READ
	})
}

// MarkAsUnread marks a message as unread.
func MarkAsUnread(ctx *WriteContext, msg *Message) error {
	return UpdateMessage(ctx, msg, map[ltp.PropID]interface{}{
		ltp.PidTagMessageFlags: int32(0),
	})
}
