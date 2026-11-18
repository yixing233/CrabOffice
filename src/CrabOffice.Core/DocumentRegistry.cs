using System;
using System.Collections.Generic;
using System.Linq;

namespace CrabOffice.Core
{
    /// <summary>Single write entry point for host events; Reconcile fixes drift. Not thread-safe: STA-only.</summary>
    public sealed class DocumentRegistry
    {
        private readonly Dictionary<Guid, DocumentRecord> _docs = new Dictionary<Guid, DocumentRecord>();
        private readonly Dictionary<Guid, WindowRecord> _windows = new Dictionary<Guid, WindowRecord>();

        public event Action RegistryChanged;

        public IReadOnlyCollection<DocumentRecord> Documents { get { return _docs.Values; } }
        public IReadOnlyCollection<WindowRecord> Windows { get { return _windows.Values; } }
        public Guid? ActiveWindowId { get; private set; }

        public DocumentRecord GetDoc(Guid id) { DocumentRecord d; return _docs.TryGetValue(id, out d) ? d : null; }
        public WindowRecord GetWin(Guid id) { WindowRecord w; return _windows.TryGetValue(id, out w) ? w : null; }

        public void ApplyHostEvent(DomainEvent e)
        {
            bool changed = true;
            switch (e.Kind)
            {
                case EventType.DocOpened:
                    if (!_docs.ContainsKey(e.DocumentId))
                        _docs[e.DocumentId] = new DocumentRecord(e.DocumentId, e.DisplayName, e.Path,
                            e.Path != null ? SavedState.Saved : SavedState.NeverSaved, e.IsReadOnly);
                    goto case EventType.WindowCreated;
                case EventType.WindowCreated:
                    if (!_windows.ContainsKey(e.WindowId))
                        _windows[e.WindowId] = new WindowRecord(e.WindowId, e.DocumentId);
                    break;
                case EventType.WindowDestroyed:
                    _windows.Remove(e.WindowId);
                    if (ActiveWindowId == e.WindowId) ActiveWindowId = null;
                    break;
                case EventType.DocClosedConfirmed:
                    foreach (var w in _windows.Values.Where(w => w.DocumentId == e.DocumentId).ToList())
                        _windows.Remove(w.WindowId);
                    _docs.Remove(e.DocumentId);
                    break;
                case EventType.DocSaved:
                    { var d = GetDoc(e.DocumentId); if (d != null) d.MarkSaved(e.Path); break; }
                case EventType.DocDirty:
                    { var d = GetDoc(e.DocumentId); if (d != null) d.MarkDirty(); break; }
                case EventType.WindowActivated:
                    { SetActive(e.WindowId); break; }
                case EventType.HostShutdown:
                    _docs.Clear(); _windows.Clear(); ActiveWindowId = null; break;
                default: changed = false; break;
            }
            if (changed && RegistryChanged != null) RegistryChanged();
        }

        private void SetActive(Guid windowId)
        {
            foreach (var w in _windows.Values) w.IsActive = (w.WindowId == windowId);
            ActiveWindowId = windowId;
        }

        /// <summary>Fix drift between live Word state and this registry. Idempotent.</summary>
        public ReconcileReport Reconcile(HostSnapshot snap)
        {
            var report = new ReconcileReport();
            var liveDocs = snap.Documents.ToDictionary(d => d.DocumentId);
            var liveWins = snap.Windows.ToDictionary(w => w.WindowId);

            // remove stale
            foreach (var id in _docs.Keys.Where(id => !liveDocs.ContainsKey(id)).ToList())
            {
                _docs.Remove(id); report.RemovedDocuments.Add(id);
            }
            foreach (var id in _windows.Keys.Where(id => !liveWins.ContainsKey(id)).ToList())
            {
                _windows.Remove(id); report.RemovedWindows.Add(id);
            }

            // upsert live
            foreach (var d in snap.Documents)
            {
                var rec = GetDoc(d.DocumentId);
                if (rec == null)
                {
                    _docs[d.DocumentId] = new DocumentRecord(d.DocumentId, d.DisplayName, d.Path,
                        d.Path != null ? (d.IsDirty ? SavedState.Dirty : SavedState.Saved) : SavedState.NeverSaved, d.IsReadOnly);
                    report.AddedDocuments.Add(d.DocumentId);
                }
                else
                {
                    rec.SavedState = d.Path != null ? (d.IsDirty ? SavedState.Dirty : SavedState.Saved) : SavedState.NeverSaved;
                    rec.IsReadOnly = d.IsReadOnly;
                }
            }
            foreach (var w in snap.Windows)
            {
                var rec = GetWin(w.WindowId);
                if (rec == null) { _windows[w.WindowId] = new WindowRecord(w.WindowId, w.DocumentId, w.IsActive, w.State); report.AddedWindows.Add(w.WindowId); }
                else { rec.IsActive = w.IsActive; rec.State = w.State; }
                if (w.IsActive) ActiveWindowId = w.WindowId;
            }

            // drop documents with no remaining window (Word document without window is invisible)
            foreach (var d in _docs.Values.Where(x => !_windows.Values.Any(w => w.DocumentId == x.DocumentId)).ToList())
            {
                _docs.Remove(d.DocumentId); report.RemovedDocuments.Add(d.DocumentId);
            }
            if (report.HasChanges && RegistryChanged != null) RegistryChanged();
            return report;
        }
    }

    public sealed class ReconcileReport
    {
        public List<Guid> AddedDocuments = new List<Guid>();
        public List<Guid> RemovedDocuments = new List<Guid>();
        public List<Guid> AddedWindows = new List<Guid>();
        public List<Guid> RemovedWindows = new List<Guid>();
        public bool HasChanges { get { return AddedDocuments.Count + RemovedDocuments.Count + AddedWindows.Count + RemovedWindows.Count > 0; } }
    }
}
