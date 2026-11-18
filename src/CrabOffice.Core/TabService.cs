using System;
using System.Collections.Generic;
using System.Linq;

namespace CrabOffice.Core
{
    /// <summary>Tab UI logic: order, pinned, activate. v1: one tab per Word window.</summary>
    public sealed class TabService
    {
        private readonly DocumentRegistry _registry;
        private readonly IWordHostAdapter _host;
        private readonly List<Guid> _orderedTabIds = new List<Guid>(); // TabId == WindowId in v1
        private readonly List<Guid> _pinned = new List<Guid>();

        public event Action TabsChanged;

        public TabService(DocumentRegistry registry, IWordHostAdapter host)
        {
            if (registry == null) throw new ArgumentNullException("registry");
            if (host == null) throw new ArgumentNullException("host");
            _registry = registry;
            _host = host;
            _registry.RegistryChanged += () => { RebuildOrders(); if (TabsChanged != null) TabsChanged(); };
        }

        public IReadOnlyList<TabRecord> GetTabs()
        {
            var list = new List<TabRecord>();
            int i = 0;
            foreach (var w in _registry.Windows.OrderBy(w => _orderedTabIds.IndexOf(w.WindowId) >= 0 ? _orderedTabIds.IndexOf(w.WindowId) : int.MaxValue))
            {
                bool pinned = _pinned.Contains(w.WindowId);
                list.Add(new TabRecord(w.WindowId, w.WindowId, w.DocumentId, i++, pinned));
            }
            return list;
        }

        public TabRecord GetActiveTab()
        {
            var active = _registry.ActiveWindowId;
            if (active == Guid.Empty) return null;
            return GetTabs().FirstOrDefault(t => t.WindowId == active);
        }

        public void ActivateTab(Guid tabId)
        {
            if (_registry.GetWin(tabId) == null) return; // stale tab, ignore (reconcile will clean)
            _host.ActivateWindow(tabId);
        }

        public void MoveTab(Guid tabId, int newOrder)
        {
            _orderedTabIds.Remove(tabId);
            _orderedTabIds.Insert(Math.Max(0, Math.Min(newOrder, _orderedTabIds.Count)), tabId);
            if (TabsChanged != null) TabsChanged();
        }

        public void PinTab(Guid tabId, bool pinned)
        {
            _pinned.Remove(tabId);
            if (pinned) _pinned.Insert(0, tabId);
            if (TabsChanged != null) TabsChanged();
        }

        public void CloseTab(Guid tabId)
        {
            if (_registry.GetWin(tabId) == null) return;
            _host.CloseWindow(tabId); // Word native close flow (save prompt included); label removal on DocClosedConfirmed/reconcile
        }

        public IReadOnlyList<Guid> CloseOtherTabs(Guid keepTabId)
        {
            var targets = GetTabs().Where(t => !t.IsPinned && t.TabId != keepTabId).Select(t => t.TabId).ToList();
            foreach (var id in targets) _host.CloseWindow(id);
            return targets;
        }

        private void RebuildOrders()
        {
            var known = _registry.Windows.Select(w => w.WindowId).ToList();
            _orderedTabIds.RemoveAll(id => !known.Contains(id));
            _pinned.RemoveAll(id => !known.Contains(id));
            foreach (var id in known.Where(id => !_orderedTabIds.Contains(id))) _orderedTabIds.Add(id);
        }
    }
}
