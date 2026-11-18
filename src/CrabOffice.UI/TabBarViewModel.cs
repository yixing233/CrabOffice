using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using CrabOffice.Core;

namespace CrabOffice.UI
{
    /// <summary>Binds DocumentRegistry/TabService to the WPF tab bar. UI thread only; never touches COM.</summary>
    public sealed class TabBarViewModel
    {
        private readonly TabService _tabs;
        private readonly Func<Guid, string> _titleResolver;

        public ObservableCollection<TabItemViewModel> Items { get; } = new ObservableCollection<TabItemViewModel>();

        public TabBarViewModel(DocumentRegistry registry, TabService tabs, Func<Guid, string> titleResolver)
        {
            if (registry == null) throw new ArgumentNullException("registry");
            if (tabs == null) throw new ArgumentNullException("tabs");
            if (titleResolver == null) throw new ArgumentNullException("titleResolver");
            _tabs = tabs;
            _titleResolver = titleResolver;
            registry.RegistryChanged += Refresh;
            tabs.TabsChanged += Refresh;
            Refresh();
        }

        public void Refresh()
        {
            var live = _tabs.GetTabs().ToList();
            var active = _tabs.GetActiveTab();
            Items.Clear();
            foreach (var t in live)
                Items.Add(new TabItemViewModel(t.TabId)
                {
                    Title = _titleResolver(t.DocumentId),
                    IsPinned = t.IsPinned,
                    IsActive = active != null && active.TabId == t.TabId
                });
        }

        /// <summary>Host layer pushes dirty/readonly/title deltas so UI never touches Word objects.</summary>
        public void UpdateItem(Guid tabId, string title, bool? isDirty, bool? isActive, bool? isReadOnly)
        {
            var item = Items.FirstOrDefault(i => i.TabId == tabId);
            if (item == null) return;
            if (title != null) item.Title = title;
            if (isDirty.HasValue) item.IsDirty = isDirty.Value;
            if (isActive.HasValue) item.IsActive = isActive.Value;
            if (isReadOnly.HasValue) item.IsReadOnly = isReadOnly.Value;
        }

        public void Activate(Guid tabId) { _tabs.ActivateTab(tabId); }
        public void Close(Guid tabId) { _tabs.CloseTab(tabId); }
    }

    public sealed class TabItemViewModel : INotifyPropertyChanged
    {
        private string _title;
        private bool _isDirty, _isActive, _isReadOnly, _isPinned;

        public Guid TabId { get; }
        public string Title { get { return _title; } set { _title = value; Raise("Title"); } }
        public bool IsDirty { get { return _isDirty; } set { _isDirty = value; Raise("IsDirty"); } }
        public bool IsActive { get { return _isActive; } set { _isActive = value; Raise("IsActive"); } }
        public bool IsReadOnly { get { return _isReadOnly; } set { _isReadOnly = value; Raise("IsReadOnly"); } }
        public bool IsPinned { get { return _isPinned; } set { _isPinned = value; Raise("IsPinned"); } }

        public TabItemViewModel(Guid tabId) { TabId = tabId; }
        private void Raise(string p) { if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(p)); }
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
