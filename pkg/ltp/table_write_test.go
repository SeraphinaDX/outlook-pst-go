package ltp

import (
	"bytes"
	"strings"
	"testing"

	"github.com/grokify/outlook-pst-go/pkg/disk"
)

func TestTableWriterBuildIsRepeatable(t *testing.T) {
	w := CreateContentsTable(disk.FormatUnicode)
	rowID, err := w.AddRowWithID(0x424)
	if err != nil {
		t.Fatal(err)
	}
	if err := w.SetRowString(rowID, PidTagSubject, strings.Repeat("subject ", 150)); err != nil {
		t.Fatal(err)
	}
	if err := w.SetRowString(rowID, PidTagMessageClass, "IPM.Note"); err != nil {
		t.Fatal(err)
	}

	first, err := w.Build()
	if err != nil {
		t.Fatal(err)
	}
	second, err := w.Build()
	if err != nil {
		t.Fatal(err)
	}

	if !bytes.Equal(first, second) {
		t.Fatalf("repeated Build changed table data: first=%d bytes second=%d bytes", len(first), len(second))
	}
}

func TestHeapWriterSpansPages(t *testing.T) {
	h := CreateTableContextHeap(disk.FormatUnicode)
	var lastPage uint16
	for i := 0; i < 12; i++ {
		hid, err := h.Allocate(bytes.Repeat([]byte{byte(i)}, 1000))
		if err != nil {
			t.Fatal(err)
		}
		lastPage = hid.PageIndex()
	}
	if lastPage == 0 {
		t.Fatal("expected allocations to span multiple heap pages")
	}

	data, err := h.Build()
	if err != nil {
		t.Fatal(err)
	}
	if len(data) <= disk.MaxDataBlockSizeUnicode {
		t.Fatalf("multi-page heap serialized to only %d bytes", len(data))
	}
	if got := len(data[:disk.MaxDataBlockSizeUnicode]); got != disk.MaxDataBlockSizeUnicode {
		t.Fatalf("first heap page = %d bytes", got)
	}
}
