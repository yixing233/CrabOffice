using System;
using System.Collections.Generic;
using CrabOffice.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CrabOffice.WordHost
{
    /// <summary>
    /// Bridges Word COM events to domain events and executes tab commands.
    /// All members run on the Word STA thread. Never persist COM references here.
    /// </summary>
    public sealed class WordHostAdapter : IWordHostAdapter
    {
        private readonly ThisAddIn _addIn;
        private readonly Word.Application _app;
        private readonly Dictionary<Guid, Word.Document> _docMap = new Dictionary<Guid, Word.Document>();
        private readonly Dictionary<Guid, Word.Window> _winMap = new Dictionary<Guid, Word.Window>();
        private readonly Dictionary<object, Guid> _docRev = new Dictionary<object, Guid>();   // RCW fast path (RCW identity is NOT reliable across acquisitions)
        private readonly Dictionary<object, Guid> _winRev = new Dictionary<object, Guid>();
        private readonly Dictionary<string, Guid> _docKeyMap = new Dictionary<string, Guid>(); // identityKey -> docId (A13: path/Name is the merge key)
        private readonly Dictionary<Guid, string> _docKeyById = new Dictionary<Guid, string>();

        /// <summary>Stable per-session identity: canonical path for saved docs, unique Name for unsaved ones.</summary>
        private static string IdentityKey(Word.Document doc)
        {
            try
            {
                if (!string.IsNullOrEmpty(doc.Path)) return doc.FullName.ToLowerInvariant();
                return doc.Name.ToLowerInvariant();
            }
            catch { return null; }
        }

        /// <summary>Re-map identity after Save/SaveAs changed the key; keeps the same docId.</summary>
        private void ReKeyDoc(Word.Document doc, Guid docId)
        {
            if (doc == null || docId == Guid.Empty) return;
            var key = IdentityKey(doc);
            if (key == null) return;
            string oldKey;
            if (_docKeyById.TryGetValue(docId, out oldKey))
            {
                if (oldKey == key) return;
                _docKeyMap.Remove(oldKey);
            }
            _docKeyMap[key] = docId;
            _docKeyById[docId] = key;
            ThisAddIn.SafeLog("rekey doc=" + Short(docId) + " -> " + key);
        }
        private readonly Guid _sessionId = Guid.NewGuid();

        public event Action<DomainEvent> HostEvent;

        public WordHostAdapter(ThisAddIn addIn)
        {
            _addIn = addIn;
            _app = addIn.Application;
            _app.DocumentOpen += OnDocumentOpen;
            ((Word.ApplicationEvents4_Event)_app).NewDocument += OnDocumentOpen; // NewDocument is ambiguous between interfaces
            _app.DocumentBeforeClose += OnDocumentBeforeClose;
            _app.DocumentBeforeSave += OnDocumentBeforeSave;
            _app.WindowActivate += OnWindowActivate;
            _app.WindowDeactivate += OnWindowDeactivate;
        }

        // ---- identity assignment ----
        private Guid DocId(Word.Document doc)
        {
            if (doc == null) return Guid.Empty;
            Guid g;
            if (_docRev.TryGetValue(doc, out g)) { _docMap[g] = doc; return g; }
            var key = IdentityKey(doc);
            if (key != null && _docKeyMap.TryGetValue(key, out g))
            {
                // same document reached via a different RCW (multi-window / reacquired) — merge identities (A13 fix)
                _docRev[doc] = g; _docMap[g] = doc;
                return g;
            }
            g = Guid.NewGuid();
            _docRev[doc] = g; _docMap[g] = doc;
            if (key != null) { _docKeyMap[key] = g; _docKeyById[g] = key; }
            return g;
        }
        private Guid WinId(Word.Window win)
        {
            if (win == null) return Guid.Empty;
            Guid g;
            if (!_winRev.TryGetValue(win, out g))
            {
                g = Guid.NewGuid();
                _winRev[win] = g; _winMap[g] = win;
            }
            return g;
        }

        private void Emit(EventType kind, Guid docId, Guid winId, string name = null, string path = null, SavedState? st = null, bool readOnly = false)
        {
            ThisAddIn.SafeLog(kind + " doc=" + Short(docId) + " win=" + Short(winId) + " name=" + (name ?? "-"));
            var h = HostEvent; if (h != null) h(new DomainEvent(kind, docId, winId, name, path, st, readOnly));
        }
        private static string Short(Guid g) { return g == Guid.Empty ? "-" : g.ToString("N").Substring(0, 8); }

        // ---- Word events (A02) ----
        private void OnDocumentOpen(Word.Document doc)
        {
            var docId = DocId(doc);
            ReKeyDoc(doc, docId);
            var win = ActiveWindowSafe();
            var winId = win != null ? WinId(win) : Guid.Empty;
            Emit(EventType.DocOpened, docId, winId, doc.Name, PathSafe(doc), null, doc.ReadOnly);
            if (winId != Guid.Empty) Emit(EventType.WindowActivated, docId, winId);
            if (winId != Guid.Empty && ThisAddIn.Presenter != null) ThisAddIn.Presenter.NotifyNewWindow(winId);
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            // NOTE: this is "about to close" (may be cancelled by save prompt).
            // Tab removal happens ONLY via reconcile after the window really dies (A06).
            ThisAddIn.SafeLog("beforeClose doc=" + Short(DocId(doc)) + " cancel=" + cancel);
        }

        private void OnDocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            // saved-state confirmation arrives via reconcile after save completes
            var docId = DocId(doc);
            ReKeyDoc(doc, docId);
            Emit(EventType.DocSaved, docId, Guid.Empty, doc.Name, PathSafe(doc));
        }

        private void OnWindowActivate(Word.Document doc, Word.Window win)
        {
            var docId = DocId(doc); var winId = WinId(win);
            Emit(EventType.WindowCreated, docId, winId, doc.Name, PathSafe(doc), null, doc.ReadOnly);
            Emit(EventType.WindowActivated, docId, winId);
            SyncDirtyState(doc, docId);
        }

        private void OnWindowDeactivate(Word.Document doc, Word.Window win)
        {
            SyncDirtyState(doc, DocId(doc));
        }

        private void SyncDirtyState(Word.Document doc, Guid docId)
        {
            if (doc == null || docId == Guid.Empty) return;
            try
            {
                Emit(EventType.DocDirty, docId, Guid.Empty, doc.Name, PathSafe(doc), doc.Saved ? SavedState.Saved : SavedState.Dirty, doc.ReadOnly);
            }
            catch (Exception ex) { ThisAddIn.SafeLog("SyncDirtyState failed: " + ex.Message); }
        }

        private static string PathSafe(Word.Document doc)
        {
            try { return string.IsNullOrEmpty(doc.Path) ? null : doc.FullName; }
            catch { return null; }
        }
        private Word.Window ActiveWindowSafe()
        {
            try { return _app.ActiveWindow; } catch { return null; }
        }

        // ---- IWordHostAdapter ----
        public HostSnapshot EnumerateSnapshot()
        {
            var snap = new HostSnapshot();
            var seenDocs = new HashSet<Guid>();
            foreach (Word.Window win in _app.Windows)
            {
                Word.Document doc = null;
                try { doc = win.Document; } catch { continue; }
                var docId = DocId(doc); var winId = WinId(win);
                if (seenDocs.Add(docId))
                    snap.Documents.Add(new HostSnapshot.DocInfo
                    {
                        DocumentId = docId, DisplayName = doc.Name, Path = PathSafe(doc),
                        IsDirty = !doc.Saved, IsReadOnly = doc.ReadOnly
                    });
                snap.Windows.Add(new HostSnapshot.WinInfo
                {
                    WindowId = winId, DocumentId = docId, IsActive = ReferenceEquals(win, ActiveWindowSafe()), State = WindowStateEx.Normal
                });
            }
            return snap;
        }

        public void ActivateWindow(Guid windowId)
        {
            Word.Window win;
            if (!_winMap.TryGetValue(windowId, out win)) { ThisAddIn.SafeLog("activate: unknown win " + Short(windowId)); return; }
            try { win.Activate(); } catch (Exception ex) { ThisAddIn.SafeLog("activate failed: " + ex.Message); }
        }

        public void CloseWindow(Guid windowId)
        {
            Word.Window win;
            if (!_winMap.TryGetValue(windowId, out win)) { ThisAddIn.SafeLog("close: unknown win " + Short(windowId)); return; }
            try
            {
                object prompt = Word.WdSaveOptions.wdPromptToSaveChanges; // native save flow (A06)
                win.Close(ref prompt);
            }
            catch (Exception ex) { ThisAddIn.SafeLog("close failed: " + ex.Message); }
        }

        /// <summary>Live window lookup for the pane presenter.</summary>
        public Word.Window GetWindow(Guid windowId)
        {
            Word.Window win; return _winMap.TryGetValue(windowId, out win) ? win : null;
        }

        public void SaveDocument(Guid documentId)
        {
            Word.Document doc;
            if (!_docMap.TryGetValue(documentId, out doc)) return;
            try { doc.Save(); } catch (Exception ex) { ThisAddIn.SafeLog("save failed: " + ex.Message); }
        }
    }
}
