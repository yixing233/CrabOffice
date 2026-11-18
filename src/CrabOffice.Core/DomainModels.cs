using System;

namespace CrabOffice.Core
{
    public enum SavedState { Saved, Dirty, NeverSaved }
    public enum WindowStateEx { Normal, Minimized, Maximized, Hidden }

    /// <summary>One Word document. DocumentId is a session GUID; path is metadata, not identity.</summary>
    public sealed class DocumentRecord
    {
        public Guid DocumentId { get; }
        public string DisplayName { get; private set; }
        public string CanonicalPath { get; private set; }
        public SavedState SavedState { get; internal set; }
        public bool IsReadOnly { get; set; }

        public DocumentRecord(Guid documentId, string displayName, string canonicalPath, SavedState savedState, bool isReadOnly = false)
        {
            DocumentId = documentId;
            DisplayName = displayName;
            CanonicalPath = canonicalPath;
            SavedState = savedState;
            IsReadOnly = isReadOnly;
        }

        public void MarkSaved(string path)
        {
            SavedState = SavedState.Saved;
            if (!string.IsNullOrEmpty(path)) { CanonicalPath = path; DisplayName = PathName(path); }
        }
        public void MarkDirty() { if (SavedState != SavedState.NeverSaved) SavedState = SavedState.Dirty; }
        internal static string PathName(string p) { var i = p != null ? p.LastIndexOf('\\') : -1; return i >= 0 ? p.Substring(i + 1) : p; }
    }

    /// <summary>One Word.Window. Multiple windows may share one DocumentId.</summary>
    public sealed class WindowRecord
    {
        public Guid WindowId { get; }
        public Guid DocumentId { get; }
        public bool IsActive { get; internal set; }
        public WindowStateEx State { get; internal set; }

        public WindowRecord(Guid windowId, Guid documentId, bool isActive = false, WindowStateEx state = WindowStateEx.Normal)
        {
            WindowId = windowId;
            DocumentId = documentId;
            IsActive = isActive;
            State = state;
        }
    }

    /// <summary>UI tab bound to a window (v1: one tab per window).</summary>
    public sealed class TabRecord
    {
        public Guid TabId { get; }
        public Guid WindowId { get; }
        public Guid DocumentId { get; }
        public int Order { get; internal set; }
        public bool IsPinned { get; internal set; }

        public TabRecord(Guid tabId, Guid windowId, Guid documentId, int order, bool isPinned = false)
        {
            TabId = tabId; WindowId = windowId; DocumentId = documentId; Order = order; IsPinned = isPinned;
        }
    }
}
