using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using CrabOffice.Core;
using CrabOffice.UI;

namespace CrabOffice.Demo
{
    /// <summary>Fake-host harness: exercises the real TabService + TabBarControl with simulated events.</summary>
    public partial class MainWindow : Window
    {
        private readonly FakeWordHost _host = new FakeWordHost();
        private readonly DocumentRegistry _registry = new DocumentRegistry();
        private TabService _tabs;
        private TabBarViewModel _vm;
        private int _docCounter;

        public MainWindow()
        {
            InitializeComponent();
            _tabs = new TabService(_registry, _host);
            _vm = new TabBarViewModel(_registry, _tabs, id =>
            {
                var d = _registry.GetDoc(id);
                return d != null ? d.DisplayName : "?";
            });
            _host.EventSink = e => _registry.ApplyHostEvent(e);
            TabBar.DataContext = _vm;
            LogLine("harness ready - click the buttons to drive the real tab bar");
        }

        private void LogLine(string s)
        {
            Log.Text = DateTime.Now.ToString("HH:mm:ss") + "  " + s + Environment.NewLine + Log.Text;
        }

        private void OnOpenDoc(object sender, RoutedEventArgs e)
        {
            _docCounter++;
            var docId = Guid.NewGuid(); var winId = Guid.NewGuid();
            var name = "Doc" + _docCounter + ".docx";
            _host.Register(docId, winId, name, _docCounter % 2 == 1);
            _host.EmitOpen(docId, winId, name);
            LogLine("opened " + name);
        }

        private void OnMakeDirty(object sender, RoutedEventArgs e)
        {
            var active = _tabs.GetActiveTab();
            if (active == null) { LogLine("no active tab"); return; }
            _host.EmitDirty(active.DocumentId);
            LogLine("dirty: " + _registry.GetDoc(active.DocumentId).DisplayName);
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            var active = _tabs.GetActiveTab();
            if (active == null) { LogLine("no active tab"); return; }
            var doc = _registry.GetDoc(active.DocumentId);
            _host.EmitSaved(active.DocumentId, "C:/demo/" + doc.DisplayName);
            LogLine("saved: " + doc.DisplayName + " -> state=" + doc.SavedState);
        }

        private void OnReconcile(object sender, RoutedEventArgs e)
        {
            var report = _registry.Reconcile(_host.BuildSnapshot());
            LogLine("reconcile: +" + report.AddedDocuments.Count + " -" + report.RemovedDocuments.Count
                + " docs, +" + report.AddedWindows.Count + " -" + report.RemovedWindows.Count + " windows");
        }
    }

    /// <summary>Stands in for IWordHostAdapter; emits DomainEvents like the real Word adapter will.</summary>
    public class FakeWordHost : IWordHostAdapter
    {
        private readonly Dictionary<Guid, Tuple<Guid, string, string>> _docs = new Dictionary<Guid, Tuple<Guid, string, string>>();
        private readonly Dictionary<Guid, Tuple<Guid, bool>> _wins = new Dictionary<Guid, Tuple<Guid, bool>>();
        private readonly List<Guid> _order = new List<Guid>();
        private Action<DomainEvent> _sink;
        public Action<DomainEvent> EventSink { set { _sink = value; } }
        public event Action<DomainEvent> HostEvent { add { } remove { } }

        public void Register(Guid docId, Guid winId, string name, bool readOnly)
        {
            _docs[docId] = Tuple.Create(winId, name, "C:/demo/" + name);
            _wins[winId] = Tuple.Create(docId, false);
            _order.Add(winId);
        }
        public void EmitOpen(Guid docId, Guid winId, string name)
        {
            _sink(new DomainEvent(EventType.DocOpened, docId, winId, name, "C:/demo/" + name));
            _sink(new DomainEvent(EventType.WindowActivated, docId, winId));
        }
        public void EmitDirty(Guid docId) { _sink(new DomainEvent(EventType.DocDirty, docId)); }
        public void EmitSaved(Guid docId, string path) { _sink(new DomainEvent(EventType.DocSaved, docId, path: path)); }

        public HostSnapshot EnumerateSnapshot() { return BuildSnapshot(); }
        public HostSnapshot BuildSnapshot()
        {
            var snap = new HostSnapshot();
            foreach (var d in _docs) snap.Documents.Add(new HostSnapshot.DocInfo { DocumentId = d.Key, DisplayName = d.Value.Item2, Path = d.Value.Item3 });
            Guid lastWin = _order.LastOrDefault();
            foreach (var w in _wins) snap.Windows.Add(new HostSnapshot.WinInfo { WindowId = w.Key, DocumentId = w.Value.Item1, IsActive = (w.Key == lastWin), State = WindowStateEx.Normal });
            return snap;
        }
        public void ActivateWindow(Guid windowId)
        {
            if (!_wins.ContainsKey(windowId)) return;
            for (int i = 0; i < _order.Count; i++) _wins[_order[i]] = Tuple.Create(_wins[_order[i]].Item1, _order[i] == windowId);
            _sink(new DomainEvent(EventType.WindowActivated, _wins[windowId].Item1, windowId));
        }
        public void CloseWindow(Guid windowId)
        {
            if (!_wins.ContainsKey(windowId)) return;
            var docId = _wins[windowId].Item1;
            _wins.Remove(windowId); _docs.Remove(docId); _order.Remove(windowId);
            _sink(new DomainEvent(EventType.DocClosedConfirmed, docId, windowId));
        }
        public void SaveDocument(Guid documentId) { EmitSaved(documentId, _docs.ContainsKey(documentId) ? _docs[documentId].Item3 : null); }
    }
}
