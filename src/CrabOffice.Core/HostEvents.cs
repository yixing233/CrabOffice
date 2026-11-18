using System;
using System.Collections.Generic;

namespace CrabOffice.Core
{
    public enum EventType
    {
        DocOpened, DocClosedConfirmed, DocSaved, DocDirty,
        WindowCreated, WindowDestroyed, WindowActivated, HostShutdown
    }

    public sealed class DomainEvent
    {
        public EventType Kind;
        public Guid DocumentId;
        public Guid WindowId;
        public string DisplayName;
        public string Path;
        public SavedState? SavedState;
        public bool IsReadOnly;

        public DomainEvent(EventType kind, Guid documentId, Guid windowId = default(Guid),
            string displayName = null, string path = null, SavedState? savedState = null, bool isReadOnly = false)
        {
            Kind = kind; DocumentId = documentId; WindowId = windowId;
            DisplayName = displayName; Path = path; SavedState = savedState; IsReadOnly = isReadOnly;
        }
    }

    /// <summary>Snapshot of live Word state, taken for startup init and Reconcile.</summary>
    public sealed class HostSnapshot
    {
        public sealed class DocInfo { public Guid DocumentId; public string DisplayName; public string Path; public bool IsDirty; public bool IsReadOnly; }
        public sealed class WinInfo { public Guid WindowId; public Guid DocumentId; public bool IsActive; public WindowStateEx State; }
        public List<DocInfo> Documents = new List<DocInfo>();
        public List<WinInfo> Windows = new List<WinInfo>();
    }

    /// <summary>Host layer contract. Implementations must invoke all members on the Word STA thread.</summary>
    public interface IWordHostAdapter
    {
        HostSnapshot EnumerateSnapshot();
        void ActivateWindow(Guid windowId);
        void CloseWindow(Guid windowId);      // delegates to Word native close flow
        void SaveDocument(Guid documentId);
        event Action<DomainEvent> HostEvent;
    }
}
