using System;
using System.Linq;
using CrabOffice.Core;
using Xunit;

namespace CrabOffice.Core.Tests
{
    // Deterministic GUIDs for reproducible tests
    internal static class G
    {
        public static Guid D(int i) => new Guid("D0000000-0000-0000-0000-0000000000" + i.ToString("D2"));
        public static Guid W(int i) => new Guid("A0000000-0000-0000-0000-0000000000" + i.ToString("D2"));
    }

    public class FakeHost : IWordHostAdapter
    {
        public HostSnapshot Next = new HostSnapshot();
        public Guid? LastActivated; public Guid? LastClosed;
        public event Action<DomainEvent> HostEvent;
        public void Emit(DomainEvent e) => HostEvent?.Invoke(e);
        public HostSnapshot EnumerateSnapshot() => Next;
        public void ActivateWindow(Guid windowId) => LastActivated = windowId;
        public void CloseWindow(Guid windowId) => LastClosed = windowId;
        public void SaveDocument(Guid documentId) { }
    }

    public class RegistryTests
    {
        [Fact]
        public void DocOpened_CreatesDocWindowAndTab()
        {
            var reg = new DocumentRegistry();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            Assert.Single(reg.Documents);
            Assert.Single(reg.Windows);
        }

        [Fact]
        public void SameDocument_NewWindow_TwoWindowsOneDoc()
        {
            var reg = new DocumentRegistry();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.WindowCreated, G.D(1), G.W(2)));
            Assert.Equal(2, reg.Windows.Count(w => w.DocumentId == G.D(1)));
        }

        [Fact]
        public void DocClosedConfirmed_RemovesAllWindowsOfDoc()
        {
            var reg = new DocumentRegistry();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.WindowCreated, G.D(1), G.W(2)));
            reg.ApplyHostEvent(new DomainEvent(EventType.DocClosedConfirmed, G.D(1)));
            Assert.Empty(reg.Documents); Assert.Empty(reg.Windows);
        }

        [Fact]
        public void DirtyThenSaved_UpdatesState()
        {
            var reg = new DocumentRegistry();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.DocDirty, G.D(1)));
            Assert.Equal(SavedState.Dirty, reg.GetDoc(G.D(1)).SavedState);
            reg.ApplyHostEvent(new DomainEvent(EventType.DocSaved, G.D(1), path: @"C:\tmp\a.docx"));
            Assert.Equal(SavedState.Saved, reg.GetDoc(G.D(1)).SavedState);
        }

        [Fact]
        public void Reconcile_FixesDrift_RemovesStaleWindow()
        {
            var reg = new DocumentRegistry();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            var snap = new HostSnapshot();
            snap.Documents.Add(new HostSnapshot.DocInfo { DocumentId = G.D(1), DisplayName = "a.docx", Path = @"C:\tmp\a.docx" });
            // live has no windows: registry must drop the stale window (and doc becomes windowless -> removed)
            var report = reg.Reconcile(snap);
            Assert.Empty(reg.Windows); Assert.True(report.HasChanges);
        }

        [Fact]
        public void Reconcile_AddsMissingLiveDoc()
        {
            var reg = new DocumentRegistry();
            var snap = new HostSnapshot();
            snap.Documents.Add(new HostSnapshot.DocInfo { DocumentId = G.D(9), DisplayName = "b.docx", Path = @"C:\tmp\b.docx" });
            snap.Windows.Add(new HostSnapshot.WinInfo { WindowId = G.W(9), DocumentId = G.D(9), IsActive = true });
            reg.Reconcile(snap);
            Assert.Single(reg.Documents); Assert.Equal(G.W(9), reg.ActiveWindowId);
        }
    }

    public class TabServiceTests
    {
        [Fact]
        public void ActivateTab_ForwardsToHost()
        {
            var reg = new DocumentRegistry();
            var host = new FakeHost();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            var tabs = new TabService(reg, host);
            tabs.ActivateTab(G.W(1));
            Assert.Equal(G.W(1), host.LastActivated);
        }

        [Fact]
        public void ActivateTab_StaleTab_NoHostCall()
        {
            var reg = new DocumentRegistry();
            var host = new FakeHost();
            var tabs = new TabService(reg, host);
            tabs.ActivateTab(G.W(1)); // not registered
            Assert.Null(host.LastActivated);
        }

        [Fact]
        public void CloseOtherTabs_SkipsPinnedAndSelf()
        {
            var reg = new DocumentRegistry();
            var host = new FakeHost();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(2), G.W(2), "b.docx", @"C:\tmp\b.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(3), G.W(3), "c.docx", @"C:\tmp\c.docx"));
            var tabs = new TabService(reg, host);
            tabs.PinTab(G.W(1), true);
            var targets = tabs.CloseOtherTabs(G.W(2));
            Assert.Equal(new[] { G.W(3) }, targets);
        }

        [Fact]
        public void MoveTab_Reorders()
        {
            var reg = new DocumentRegistry();
            var host = new FakeHost();
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(1), G.W(1), "a.docx", @"C:\tmp\a.docx"));
            reg.ApplyHostEvent(new DomainEvent(EventType.DocOpened, G.D(2), G.W(2), "b.docx", @"C:\tmp\b.docx"));
            var tabs = new TabService(reg, host);
            tabs.MoveTab(G.W(2), 0);
            var order = tabs.GetTabs().Select(t => t.TabId).ToList();
            Assert.Equal(new[] { G.W(2), G.W(1) }, order);
        }
    }
}
