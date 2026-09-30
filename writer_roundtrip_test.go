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
			wantAttachment := []byte("attachment bytes survive the PST round trip")
			_, err = ctx.CreateMessage(inbox).
				SetSubject("MailSalonTools round trip").
				SetBody("plain body").
				SetHTMLBody("<p>html body</p>").
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
			if err != nil || html != "<p>html body</p>" {
				t.Fatalf("html = %q, %v", html, err)
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

func TestWriterLargeValuesRoundTrip(t *testing.T) {
	path := filepath.Join(t.TempDir(), "large-values.pst")
	pst, err := Create(path, disk.FormatUnicode)
	if err != nil {
		t.Fatalf("Create: %v", err)
	}

	root, err := pst.RootFolder()
	if err != nil {
		t.Fatalf("RootFolder: %v", err)
	}
	ctx, err := pst.BeginWrite()
	if err != nil {
		t.Fatalf("BeginWrite: %v", err)
	}
	folder, err := ctx.CreateFolder(root, "Large")
	if err != nil {
		t.Fatalf("CreateFolder: %v", err)
	}

	html := "<html><body>" + strings.Repeat("<p>large HTML body</p>", 2500) + "</body></html>"
	attachment := bytes.Repeat([]byte{0x00, 0x01, 0x02, 0x03, 0xFE, 0xFF}, 20000)

	_, err = ctx.CreateMessage(folder).
		SetSubject("large values").
		SetBody("plain body").
		SetHTMLBody(html).
		AddAttachmentWithMime("large.bin", attachment, "application/octet-stream").
		Build()
	if err != nil {
		t.Fatalf("Build: %v", err)
	}
	if err := ctx.Commit(); err != nil {
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
	folder, err = root.FindSubfolder("Large")
	if err != nil {
		t.Fatalf("FindSubfolder: %v", err)
	}

	var got *Message
	for msg, iterErr := range folder.Messages() {
		if iterErr != nil {
			t.Fatalf("Messages: %v", iterErr)
		}
		got = msg
		break
	}
	if got == nil {
		t.Fatal("message missing after reopen")
	}

	gotHTML, err := got.HTMLBody()
	if err != nil {
		t.Fatalf("HTMLBody: %v", err)
	}
	if gotHTML != html {
		t.Fatalf("HTML body length = %d, want %d", len(gotHTML), len(html))
	}

	var gotAttachment []byte
	for a, iterErr := range got.Attachments() {
		if iterErr != nil {
			t.Fatalf("Attachments: %v", iterErr)
		}
		gotAttachment, err = a.Data()
		if err != nil {
			t.Fatalf("attachment Data: %v", err)
		}
		break
	}
	if !bytes.Equal(gotAttachment, attachment) {
		t.Fatalf("attachment length = %d, want %d", len(gotAttachment), len(attachment))
	}
}

func TestWriterLargeContentsTableRoundTrip(t *testing.T) {
	path := filepath.Join(t.TempDir(), "large-contents.pst")
	pst, err := Create(path, disk.FormatUnicode)
	if err != nil {
		t.Fatalf("Create: %v", err)
	}

	root, err := pst.RootFolder()
	if err != nil {
		t.Fatalf("RootFolder: %v", err)
	}
	ctx, err := pst.BeginWrite()
	if err != nil {
		t.Fatalf("BeginWrite: %v", err)
	}
	folder, err := ctx.CreateFolder(root, "Bulk")
	if err != nil {
		t.Fatalf("CreateFolder: %v", err)
	}

	const wantMessages = 500
	for i := 0; i < wantMessages; i++ {
		subject := fmt.Sprintf("bulk message %04d %s", i, strings.Repeat("x", 48))
		if _, err := ctx.CreateMessage(folder).
			SetSubject(subject).
			SetBody("small body").
			Build(); err != nil {
			t.Fatalf("Build message %d: %v", i, err)
		}
	}

	if err := ctx.Commit(); err != nil {
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
	folder, err = root.FindSubfolder("Bulk")
	if err != nil {
		t.Fatalf("FindSubfolder: %v", err)
	}

	got := 0
	for _, iterErr := range folder.Messages() {
		if iterErr != nil {
			t.Fatalf("Messages: %v", iterErr)
		}
		got++
	}
	if got != wantMessages {
		t.Fatalf("message count = %d, want %d", got, wantMessages)
	}
}

func TestWriterBatchedTransactionsRoundTrip(t *testing.T) {
	path := filepath.Join(t.TempDir(), "batched.pst")
	pst, err := Create(path, disk.FormatUnicode)
	if err != nil {
		t.Fatalf("Create: %v", err)
	}

	root, err := pst.RootFolder()
	if err != nil {
		t.Fatalf("RootFolder: %v", err)
	}
	ctx, err := pst.BeginWrite()
	if err != nil {
		t.Fatalf("BeginWrite: %v", err)
	}
	folder, err := ctx.CreateFolder(root, "Imported")
	if err != nil {
		t.Fatalf("CreateFolder: %v", err)
	}

	const batches = 5
	const perBatch = 100
	for batch := 0; batch < batches; batch++ {
		if batch > 0 {
			pst, err = OpenReadWrite(path)
			if err != nil {
				t.Fatalf("OpenReadWrite batch %d: %v", batch, err)
			}
			root, err = pst.RootFolder()
			if err != nil {
				t.Fatalf("RootFolder batch %d: %v", batch, err)
			}
			folder, err = root.FindSubfolder("Imported")
			if err != nil {
				t.Fatalf("FindSubfolder batch %d: %v", batch, err)
			}
			ctx, err = pst.BeginWrite()
			if err != nil {
				t.Fatalf("BeginWrite batch %d: %v", batch, err)
			}
		}

		for i := 0; i < perBatch; i++ {
			n := batch*perBatch + i
			builder := ctx.CreateMessage(folder).
				SetSubject(fmt.Sprintf("batched message %04d %s", n, strings.Repeat("y", 40))).
				SetBody("small body")
			if n == 202 {
				builder.SetHTMLBody("<html><body>"+strings.Repeat("<p>large batch HTML</p>", 2500)+"</body></html>").
					AddAttachmentWithMime("batch-large.bin", bytes.Repeat([]byte{1, 2, 3, 4}, 25000), "application/octet-stream")
			}
			if _, err := builder.Build(); err != nil {
				t.Fatalf("Build batch %d message %d: %v", batch, n, err)
			}
		}

		if err := ctx.Commit(); err != nil {
			t.Fatalf("Commit batch %d: %v", batch, err)
		}
		if err := pst.Close(); err != nil {
			t.Fatalf("Close batch %d: %v", batch, err)
		}
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
	folder, err = root.FindSubfolder("Imported")
	if err != nil {
		t.Fatalf("FindSubfolder final: %v", err)
	}

	got := 0
	largeSeen := false
	for msg, iterErr := range folder.Messages() {
		if iterErr != nil {
			t.Fatalf("Messages final: %v", iterErr)
		}
		got++
		subject, _ := msg.Subject()
		if strings.HasPrefix(subject, "batched message 0202 ") {
			html, err := msg.HTMLBody()
			if err != nil || len(html) < 40000 {
				t.Fatalf("large batched HTML missing: len=%d err=%v", len(html), err)
			}
			for attachment, attachErr := range msg.Attachments() {
				if attachErr != nil {
					t.Fatalf("large batched attachment: %v", attachErr)
				}
				data, err := attachment.Data()
				if err != nil {
					t.Fatalf("large batched attachment data: %v", err)
				}
				if len(data) == 100000 {
					largeSeen = true
				}
			}
		}
	}
	if got != batches*perBatch {
		t.Fatalf("message count = %d, want %d", got, batches*perBatch)
	}
	if !largeSeen {
		t.Fatal("large value message did not survive batched transactions")
	}
}
