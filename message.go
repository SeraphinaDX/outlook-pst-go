package outlookpst

import (
	"fmt"
	"iter"
	"sync"
	"time"

	"github.com/grokify/outlook-pst-go/pkg/ltp"
	"github.com/grokify/outlook-pst-go/pkg/ndb"
	"github.com/grokify/outlook-pst-go/pkg/rtf"
	"github.com/grokify/outlook-pst-go/pkg/util"
)

// Message represents an email message in a PST file.
type Message struct {
	pst  *PST
	node *ndb.Node
	nid  util.NodeID
	bag  *ltp.PropertyBag

	// Attachment table (lazy loaded)
	attachmentOnce  sync.Once
	attachmentTable *ltp.Table
	attachmentErr   error

	// Recipient table (lazy loaded)
	recipientOnce  sync.Once
	recipientTable *ltp.Table
	recipientErr   error
}

// newMessage creates a new Message from a node.
func newMessage(pst *PST, node *ndb.Node) (*Message, error) {
	bag, err := ltp.NewPropertyBag(node)
	if err != nil {
		return nil, fmt.Errorf("failed to create property bag: %w", err)
	}

	return &Message{
		pst:  pst,
		node: node,
		nid:  node.ID(),
		bag:  bag,
	}, nil
}

// ID returns the message's node ID.
func (m *Message) ID() util.NodeID {
	if m.node != nil {
		return m.node.ID()
	}
	return m.nid
}

// Subject returns the message subject.
func (m *Message) Subject() (string, error) {
	return m.bag.GetString(ltp.PidTagSubject)
}

// NormalizedSubject returns the normalized subject (without prefix like "Re:").
func (m *Message) NormalizedSubject() (string, error) {
	return m.bag.GetString(ltp.PidTagNormalizedSubject)
}

// Body returns the plain text body.
func (m *Message) Body() (string, error) {
	return m.bag.GetString(ltp.PidTagBody)
}

// HTMLBody returns the HTML body.
func (m *Message) HTMLBody() (string, error) {
	return m.bag.GetString(ltp.PidTagHtmlBody)
}

// RTFBody returns the compressed RTF body (raw bytes).
func (m *Message) RTFBody() ([]byte, error) {
	return m.bag.GetBinary(ltp.PidTagRtfCompressed)
}

// RTFBodyDecompressed returns the decompressed RTF body as a string.
// See [MS-OXRTFCP] for the compression format.
func (m *Message) RTFBodyDecompressed() (string, error) {
	compressed, err := m.RTFBody()
	if err != nil {
		return "", err
	}

	decompressed, err := rtf.Decompress(compressed)
	if err != nil {
		return "", fmt.Errorf("failed to decompress RTF: %w", err)
	}

	return string(decompressed), nil
}

// MessageClass returns the message class (e.g., "IPM.Note").
func (m *Message) MessageClass() (string, error) {
	return m.bag.GetString(ltp.PidTagMessageClass)
}

// SenderName returns the sender's display name.
func (m *Message) SenderName() (string, error) {
	return m.bag.GetString(ltp.PidTagSenderName)
}

// SenderEmail returns the sender's email address.
func (m *Message) SenderEmail() (string, error) {
	return m.bag.GetString(ltp.PidTagSenderEmailAddress)
}

// SentRepresentingName returns the "sent representing" display name.
func (m *Message) SentRepresentingName() (string, error) {
	return m.bag.GetString(ltp.PidTagSentRepresentingName)
}

// SentRepresentingEmail returns the "sent representing" email address.
func (m *Message) SentRepresentingEmail() (string, error) {
	return m.bag.GetString(ltp.PidTagSentRepresentingEmailAddress)
}

// DisplayTo returns the To recipients as a string.
func (m *Message) DisplayTo() (string, error) {
	return m.bag.GetString(ltp.PidTagDisplayTo)
}

// DisplayCc returns the Cc recipients as a string.
func (m *Message) DisplayCc() (string, error) {
	return m.bag.GetString(ltp.PidTagDisplayCc)
}

// DisplayBcc returns the Bcc recipients as a string.
func (m *Message) DisplayBcc() (string, error) {
	return m.bag.GetString(ltp.PidTagDisplayBcc)
}

// DeliveryTime returns the message delivery time.
func (m *Message) DeliveryTime() (time.Time, error) {
	ft, err := m.bag.GetTime(ltp.PidTagMessageDeliveryTime)
	if err != nil {
		return time.Time{}, err
	}
	return ltp.FileTimeToTime(ft), nil
}

// SubmitTime returns the time the message was submitted.
func (m *Message) SubmitTime() (time.Time, error) {
	ft, err := m.bag.GetTime(ltp.PidTagClientSubmitTime)
	if err != nil {
		return time.Time{}, err
	}
	return ltp.FileTimeToTime(ft), nil
}

// CreationTime returns the message creation time.
func (m *Message) CreationTime() (time.Time, error) {
	ft, err := m.bag.GetTime(ltp.PidTagCreationTime)
	if err != nil {
		return time.Time{}, err
	}
	return ltp.FileTimeToTime(ft), nil
}

// LastModificationTime returns the last modification time.
func (m *Message) LastModificationTime() (time.Time, error) {
	ft, err := m.bag.GetTime(ltp.PidTagLastModificationTime)
	if err != nil {
		return time.Time{}, err
	}
	return ltp.FileTimeToTime(ft), nil
}

// MessageSize returns the message size in bytes.
func (m *Message) MessageSize() (int32, error) {
	return m.bag.GetInt32(ltp.PidTagMessageSize)
}

// Importance returns the message importance level.
func (m *Message) Importance() (int32, error) {
	return m.bag.GetInt32(ltp.PidTagImportance)
}

// Priority returns the message priority.
func (m *Message) Priority() (int32, error) {
	return m.bag.GetInt32(ltp.PidTagPriority)
}

// Sensitivity returns the message sensitivity.
func (m *Message) Sensitivity() (int32, error) {
	return m.bag.GetInt32(ltp.PidTagSensitivity)
}

// HasAttachments returns true if the message has attachments.
func (m *Message) HasAttachments() (bool, error) {
	return m.bag.GetBool(ltp.PidTagHasAttachments)
}

// InternetMessageID returns the internet message ID (RFC 822).
func (m *Message) InternetMessageID() (string, error) {
	return m.bag.GetString(ltp.PidTagInternetMessageId)
}

// ConversationTopic returns the conversation topic.
func (m *Message) ConversationTopic() (string, error) {
	return m.bag.GetString(ltp.PidTagConversationTopic)
}

// ConversationIndex returns the conversation index.
func (m *Message) ConversationIndex() ([]byte, error) {
	return m.bag.GetBinary(ltp.PidTagConversationIndex)
}

// PropertyBag returns the message's property bag for advanced access.
func (m *Message) PropertyBag() *ltp.PropertyBag {
	return m.bag
}

// loadAttachmentTable loads the attachment table.
func (m *Message) loadAttachmentTable() error {
	m.attachmentOnce.Do(func() {
		// Attachment table NID is message NID with type changed to attachment table
		messageNID := m.node.ID()
		attachNID := util.MakeNID(util.NIDTypeAttachmentTable, messageNID.Index())

		// The attachment table is a subnode
		attachNode, err := m.node.LookupSubnode(attachNID)
		if err != nil {
			// No attachments
			m.attachmentErr = nil
			return
		}

		m.attachmentTable, err = ltp.NewTable(attachNode)
		if err != nil {
			m.attachmentErr = fmt.Errorf("failed to create attachment table: %w", err)
			return
		}
	})
	return m.attachmentErr
}

// loadRecipientTable loads the recipient table.
func (m *Message) loadRecipientTable() error {
	m.recipientOnce.Do(func() {
		// Recipient table NID is message NID with type changed to recipient table
		messageNID := m.node.ID()
		recipNID := util.MakeNID(util.NIDTypeRecipientTable, messageNID.Index())

		// The recipient table is a subnode
		recipNode, err := m.node.LookupSubnode(recipNID)
		if err != nil {
			// No recipients (unusual but possible)
			m.recipientErr = nil
			return
		}

		m.recipientTable, err = ltp.NewTable(recipNode)
		if err != nil {
			m.recipientErr = fmt.Errorf("failed to create recipient table: %w", err)
			return
		}
	})
	return m.recipientErr
}

// Attachments returns an iterator over attachments.
func (m *Message) Attachments() iter.Seq2[*Attachment, error] {
	return func(yield func(*Attachment, error) bool) {
		if err := m.loadAttachmentTable(); err != nil {
			yield(nil, err)
			return
		}

		if m.attachmentTable == nil {
			return // No attachments
		}

		for row, err := range m.attachmentTable.Rows() {
			if err != nil {
				yield(nil, err)
				return
			}

			// Get the attachment subnode NID from the row
			nidData, err := row.GetRaw(ltp.PidTagLtpRowId)
			if err != nil {
				yield(nil, fmt.Errorf("failed to get attachment NID: %w", err))
				return
			}

			if len(nidData) < 4 {
				continue
			}

			attachNID := util.NodeID(nidData[0]) | util.NodeID(nidData[1])<<8 |
				util.NodeID(nidData[2])<<16 | util.NodeID(nidData[3])<<24

			// Get the attachment subnode
			attachNode, err := m.node.LookupSubnode(attachNID)
			if err != nil {
				yield(nil, fmt.Errorf("failed to get attachment node: %w", err))
				return
			}

			attachment, err := newAttachment(m.pst, m, attachNode)
			if err != nil {
				yield(nil, err)
				return
			}

			if !yield(attachment, nil) {
				return
			}
		}
	}
}

// AttachmentCount returns the number of attachments.
func (m *Message) AttachmentCount() (int, error) {
	if err := m.loadAttachmentTable(); err != nil {
		return 0, err
	}
	if m.attachmentTable == nil {
		return 0, nil
	}
	return m.attachmentTable.RowCount(), nil
}

// Recipients returns an iterator over recipients.
func (m *Message) Recipients() iter.Seq2[*Recipient, error] {
	return func(yield func(*Recipient, error) bool) {
		if err := m.loadRecipientTable(); err != nil {
			yield(nil, err)
			return
		}

		if m.recipientTable == nil {
			return // No recipients
		}

		for row, err := range m.recipientTable.Rows() {
			if err != nil {
				yield(nil, err)
				return
			}

			recipient := newRecipient(row)

			if !yield(recipient, nil) {
				return
			}
		}
	}
}

// RecipientCount returns the number of recipients.
func (m *Message) RecipientCount() (int, error) {
	if err := m.loadRecipientTable(); err != nil {
		return 0, err
	}
	if m.recipientTable == nil {
		return 0, nil
	}
	return m.recipientTable.RowCount(), nil
}

// AttachmentTable returns the attachment table for advanced access.
func (m *Message) AttachmentTable() (*ltp.Table, error) {
	if err := m.loadAttachmentTable(); err != nil {
		return nil, err
	}
	return m.attachmentTable, nil
}

// RecipientTable returns the recipient table for advanced access.
func (m *Message) RecipientTable() (*ltp.Table, error) {
	if err := m.loadRecipientTable(); err != nil {
		return nil, err
	}
	return m.recipientTable, nil
}

// Attachment represents an email attachment.
type Attachment struct {
	pst     *PST
	message *Message
	node    *ndb.Node
	bag     *ltp.PropertyBag
}

// newAttachment creates a new Attachment from a subnode.
func newAttachment(pst *PST, msg *Message, node *ndb.Node) (*Attachment, error) {
	bag, err := ltp.NewPropertyBag(node)
	if err != nil {
		return nil, fmt.Errorf("failed to create property bag: %w", err)
	}

	return &Attachment{
		pst:     pst,
		message: msg,
		node:    node,
		bag:     bag,
	}, nil
}

// Filename returns the attachment filename.
func (a *Attachment) Filename() (string, error) {
	// Try long filename first
	name, err := a.bag.GetString(ltp.PidTagAttachLongFilename)
	if err == nil && name != "" {
		return name, nil
	}
	// Fall back to short filename
	return a.bag.GetString(ltp.PidTagAttachFilename)
}

// Extension returns the file extension.
func (a *Attachment) Extension() (string, error) {
	return a.bag.GetString(ltp.PidTagAttachExtension)
}

// Size returns the attachment size.
func (a *Attachment) Size() (int32, error) {
	return a.bag.GetInt32(ltp.PidTagAttachSize)
}

// Method returns the attachment method.
func (a *Attachment) Method() (int32, error) {
	return a.bag.GetInt32(ltp.PidTagAttachMethod)
}

// MimeType returns the MIME type.
func (a *Attachment) MimeType() (string, error) {
	return a.bag.GetString(ltp.PidTagAttachMimeTag)
}

// ContentID returns the content ID (for inline attachments).
func (a *Attachment) ContentID() (string, error) {
	return a.bag.GetString(ltp.PidTagAttachContentId)
}

// Data returns the attachment data.
func (a *Attachment) Data() ([]byte, error) {
	return a.bag.GetBinary(ltp.PidTagAttachDataBinary)
}

// IsEmbeddedMessage returns true if this attachment is an embedded message.
func (a *Attachment) IsEmbeddedMessage() (bool, error) {
	method, err := a.Method()
	if err != nil {
		return false, err
	}
	return method == ltp.AttachMethodEmbedded, nil
}

// OpenAsMessage opens an embedded message attachment as a Message.
func (a *Attachment) OpenAsMessage() (*Message, error) {
	isEmbedded, err := a.IsEmbeddedMessage()
	if err != nil {
		return nil, err
	}
	if !isEmbedded {
		return nil, fmt.Errorf("attachment is not an embedded message")
	}

	// The embedded message is stored as a subnode
	// Find it by looking for the message subnode
	msgNID := util.MakeNID(util.NIDTypeNormalMessage, a.node.ID().Index())
	msgNode, err := a.node.LookupSubnode(msgNID)
	if err != nil {
		return nil, fmt.Errorf("failed to find embedded message: %w", err)
	}

	return newMessage(a.pst, msgNode)
}

// PropertyBag returns the attachment's property bag for advanced access.
func (a *Attachment) PropertyBag() *ltp.PropertyBag {
	return a.bag
}

// RecipientType represents the type of recipient.
type RecipientType int32

const (
	RecipientOriginator RecipientType = 0
	RecipientTo         RecipientType = 1
	RecipientCc         RecipientType = 2
	RecipientBcc        RecipientType = 3
)

func (rt RecipientType) String() string {
	switch rt {
	case RecipientOriginator:
		return "Originator"
	case RecipientTo:
		return "To"
	case RecipientCc:
		return "Cc"
	case RecipientBcc:
		return "Bcc"
	default:
		return "Unknown"
	}
}

// Recipient represents an email recipient.
type Recipient struct {
	row *ltp.TableRow
}

// newRecipient creates a new Recipient from a table row.
func newRecipient(row *ltp.TableRow) *Recipient {
	return &Recipient{row: row}
}

// Name returns the recipient's display name.
func (r *Recipient) Name() (string, error) {
	return r.row.GetString(ltp.PidTagDisplayName)
}

// Email returns the recipient's email address.
func (r *Recipient) Email() (string, error) {
	// Try SMTP address first
	email, err := r.row.GetString(ltp.PidTagSmtpAddress)
	if err == nil && email != "" {
		return email, nil
	}
	// Fall back to email address
	return r.row.GetString(ltp.PidTagEmailAddress)
}

// Type returns the recipient type (To, Cc, Bcc).
func (r *Recipient) Type() (RecipientType, error) {
	t, err := r.row.GetInt32(ltp.PidTagRecipientType)
	if err != nil {
		return 0, err
	}
	return RecipientType(t), nil
}

// AddressType returns the address type (e.g., "SMTP", "EX").
func (r *Recipient) AddressType() (string, error) {
	return r.row.GetString(ltp.PidTagAddressType)
}

// Row returns the underlying table row for advanced access.
func (r *Recipient) Row() *ltp.TableRow {
	return r.row
}
