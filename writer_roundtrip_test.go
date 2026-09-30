package outlookpst

import (
	"bytes"
	"fmt"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/grokify/outlook-pst-go/pkg/disk"
)

func TestWriterRoundTrip(t *testing.T) {
	formats := []struct {
		name   string
		format disk.PSTFormat
	}{
		{name: "unicode", format: disk.FormatUnicode},
		{name: "ansi", format: disk.FormatANSI},
	}

	for _, tc := range formats {
		t.Run(tc.name, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "roundtrip.pst")
			pst, err := Create(path, tc.format)
			if err != nil {
				t.Fatalf("Create: %v", err)
			}

			root, err := pst.RootFolder()
			if err != nil {
				_ = pst.Close()
				t.Fatalf("RootFolder: %v", err)
			}

			ctx, err := pst.BeginWrite()
			if err != nil {
				_ = pst.Close()
				t.Fatalf("BeginWrite: %v", err)
			}

			inbox, err := ctx.CreateFolder(root, "Inbox")
			if err != nil {
				_ = ctx.Rollback()
				_ = pst.Close()
				t.Fatalf("CreateFolder: %v", err)
			}

			sent := time.Date(2026, 9, 30, 12, 0, 0, 0, time.UTC)
			wantHTML := "<p>" + strings.Repeat("large html body ", 12000) + "</p>"
			wantAttachment := bytes.Repeat([]byte{0x5A}, 128*1024)
			_, err = ctx.CreateMessage(inbox).
				SetSubject("MailSalonTools round trip").
				SetBody("plain body").
				SetHTMLBody(wantHTML).
				SetFrom("Britney", "britney@example.com").
				AddTo("Recipient", "recipient@example.com").
				AddCC("Copy", "copy@example.com").
				SetSentTime(sent).
				AddAttachmentWithMime("notes.txt", wantAttachment, "text/plain").
				Build()
			if err != nil {
				_ = ctx.Rollback()
				_ = pst.Close()
				t.Fatalf("Build message: %v", err)
			}

			if err := ctx.Commit(); err != nil {
				_ = pst.Close()
				t.Fatalf("Commit: %v", err)
			}
			if err := pst.Close(); err != nil {
				t.Fatalf("Close writer: %v", err)
			}

			pst, err = Open(path)
			if err != nil {
				t.Fatalf("Open round-trip PST: %v", err)
			}
			defer func() { _ = pst.Close() }()

			root, err = pst.RootFolder()
			if err != nil {
				t.Fatalf("RootFolder after reopen: %v", err)
			}
			inbox, err = root.FindSubfolder("Inbox")
			if err != nil {
				t.Fatalf("FindSubfolder after reopen: %v", err)
			}

			var got *Message
			for msg, iterErr := range inbox.Messages() {
				if iterErr != nil {
					t.Fatalf("iterate messages: %v", iterErr)
				}
				got = msg
				break
			}
			if got == nil {
				t.Fatal("no message found after reopen")
			}

			subject, err := got.Subject()
			if err != nil || subject != "MailSalonTools round trip" {
				t.Fatalf("subject = %q, %v", subject, err)
			}
			body, err := got.Body()
			if err != nil || body != "plain body" {
				t.Fatalf("body = %q, %v", body, err)
			}
			html, err := got.HTMLBody()
			if err != nil || html != wantHTML {
				t.Fatalf("html length = %d, want %d, err=%v", len(html), len(wantHTML), err)
			}
			submitTime, err := got.SubmitTime()
			if err != nil {
				t.Fatalf("submit time: %v", err)
			}
			if !submitTime.Equal(sent) {
				t.Fatalf("submit time = %v, want %v", submitTime, sent)
			}

			var attachmentSeen bool
			for attachment, iterErr := range got.Attachments() {
				if iterErr != nil {
					t.Fatalf("iterate attachments: %v", iterErr)
				}
				name, err := attachment.Filename()
				if err != nil {
					t.Fatalf("attachment filename: %v", err)
				}
				data, err := attachment.Data()
				if err != nil {
					t.Fatalf("attachment data: %v", err)
				}
				if name == "notes.txt" && bytes.Equal(data, wantAttachment) {
					attachmentSeen = true
				}
			}
			if !attachmentSeen {
				t.Fatal("attachment did not survive round trip")
			}

			var toSeen, ccSeen bool
			for recipient, iterErr := range got.Recipients() {
				if iterErr != nil {
					t.Fatalf("iterate recipients: %v", iterErr)
				}
				email, _ := recipient.Email()
				rtype, _ := recipient.Type()
				if email == "recipient@example.com" && rtype == RecipientTo {
					toSeen = true
				}
				if email == "copy@example.com" && rtype == RecipientCc {
					ccSeen = true
				}
			}
			if !toSeen || !ccSeen {
				t.Fatalf("recipients did not survive round trip: to=%v cc=%v", toSeen, ccSeen)
			}
		})
	}
}

func TestWriterManyMessagesContentsTable(t *testing.T) {
	path := filepath.Join(t.TempDir(), "many-messages.pst")
	pst, err := Create(path, disk.FormatUnicode)
	if err != nil {
		t.Fatalf("Create: %v", err)
	}

	root, err := pst.RootFolder()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("RootFolder: %v", err)
	}
	ctx, err := pst.BeginWrite()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("BeginWrite: %v", err)
	}
	inbox, err := ctx.CreateFolder(root, "Inbox")
	if err != nil {
		_ = ctx.Rollback()
		_ = pst.Close()
		t.Fatalf("CreateFolder: %v", err)
	}

	const count = 20
	for i := 0; i < count; i++ {
		subject := strings.Repeat("long subject segment ", 60) + fmt.Sprintf(" #%02d", i)
		if _, err := ctx.CreateMessage(inbox).
			SetSubject(subject).
			SetBody("body").
			Build(); err != nil {
			_ = ctx.Rollback()
			_ = pst.Close()
			t.Fatalf("Build message %d: %v", i+1, err)
		}
	}

	if err := ctx.Commit(); err != nil {
		_ = pst.Close()
		t.Fatalf("Commit: %v", err)
	}
	if err := pst.Close(); err != nil {
		t.Fatalf("Close: %v", err)
	}

	pst, err = Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer func() { _ = pst.Close() }()

	root, err = pst.RootFolder()
	if err != nil {
		t.Fatalf("RootFolder after reopen: %v", err)
	}
	inbox, err = root.FindSubfolder("Inbox")
	if err != nil {
		t.Fatalf("FindSubfolder: %v", err)
	}

	got := 0
	for _, iterErr := range inbox.Messages() {
		if iterErr != nil {
			t.Fatalf("iterate messages: %v", iterErr)
		}
		got++
	}
	if got != count {
		t.Fatalf("message count = %d, want %d", got, count)
	}
}

func TestWriterAppendAfterReopenKeepsUniqueNIDs(t *testing.T) {
	path := filepath.Join(t.TempDir(), "append-after-reopen.pst")
	pst, err := Create(path, disk.FormatUnicode)
	if err != nil {
		t.Fatalf("Create: %v", err)
	}

	root, err := pst.RootFolder()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("RootFolder: %v", err)
	}
	ctx, err := pst.BeginWrite()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("BeginWrite: %v", err)
	}
	inbox, err := ctx.CreateFolder(root, "Inbox")
	if err != nil {
		_ = ctx.Rollback()
		_ = pst.Close()
		t.Fatalf("CreateFolder: %v", err)
	}

	const firstBatch = 100
	for i := 0; i < firstBatch; i++ {
		if _, err := ctx.CreateMessage(inbox).
			SetSubject(fmt.Sprintf("before reopen %03d", i)).
			SetBody("body").
			Build(); err != nil {
			_ = ctx.Rollback()
			_ = pst.Close()
			t.Fatalf("Build first batch message %d: %v", i+1, err)
		}
	}
	if err := ctx.Commit(); err != nil {
		_ = pst.Close()
		t.Fatalf("Commit first batch: %v", err)
	}
	if err := pst.Close(); err != nil {
		t.Fatalf("Close first batch: %v", err)
	}

	pst, err = OpenReadWrite(path)
	if err != nil {
		t.Fatalf("OpenReadWrite: %v", err)
	}
	root, err = pst.RootFolder()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("RootFolder after reopen: %v", err)
	}
	inbox, err = root.FindSubfolder("Inbox")
	if err != nil {
		_ = pst.Close()
		t.Fatalf("FindSubfolder after reopen: %v", err)
	}
	ctx, err = pst.BeginWrite()
	if err != nil {
		_ = pst.Close()
		t.Fatalf("BeginWrite after reopen: %v", err)
	}

	const secondBatch = 5
	for i := 0; i < secondBatch; i++ {
		if _, err := ctx.CreateMessage(inbox).
			SetSubject(fmt.Sprintf("after reopen %03d", i)).
			SetBody("body").
			Build(); err != nil {
			_ = ctx.Rollback()
			_ = pst.Close()
			t.Fatalf("Build second batch message %d: %v", i+1, err)
		}
	}
	if err := ctx.Commit(); err != nil {
		_ = pst.Close()
		t.Fatalf("Commit second batch: %v", err)
	}
	if err := pst.Close(); err != nil {
		t.Fatalf("Close second batch: %v", err)
	}

	pst, err = Open(path)
	if err != nil {
		t.Fatalf("Open final PST: %v", err)
	}
	defer func() { _ = pst.Close() }()

	root, err = pst.RootFolder()
	if err != nil {
		t.Fatalf("RootFolder final: %v", err)
	}
	inbox, err = root.FindSubfolder("Inbox")
	if err != nil {
		t.Fatalf("FindSubfolder final: %v", err)
	}

	got := 0
	for _, iterErr := range inbox.Messages() {
		if iterErr != nil {
			t.Fatalf("iterate final messages: %v", iterErr)
		}
		got++
	}
	want := firstBatch + secondBatch
	if got != want {
		t.Fatalf("message count = %d, want %d", got, want)
	}
}
