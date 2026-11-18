using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Controls;
using CrabOffice.UI;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace CrabOffice.WordHost
{
    /// <summary>
    /// Creates one top-docked CustomTaskPane per Word.Window hosting the shared WPF tab bar (A01/A15).
    /// All panes share one TabBarViewModel; UI never touches COM.
    /// </summary>
    public sealed class TabPanePresenter : IDisposable
    {
        private readonly ThisAddIn _addIn;
        private readonly TabBarViewModel _vm;
        private readonly Dictionary<Guid, CustomTaskPane> _panes = new Dictionary<Guid, CustomTaskPane>();
        private readonly object _paneKey = new object();

        public TabPanePresenter(ThisAddIn addIn, TabBarViewModel vm)
        {
            _addIn = addIn; _vm = vm;
        }

    /// <summary>Minimal WinForms UserControl that hosts the WPF tab bar via ElementHost.</summary>
    public class PaneHostControl : System.Windows.Forms.UserControl
    {
        public readonly ElementHost ElementHost = new ElementHost();
        public PaneHostControl()
        {
            ElementHost.Dock = DockStyle.Fill;
            Controls.Add(ElementHost);
        }
    }

        public void EnsurePane(Guid windowId)
        {
            if (_panes.ContainsKey(windowId)) return;
            try
            {
                var win = ResolveWindow(windowId);
                if (win == null) { ThisAddIn.SafeLog("EnsurePane: cannot resolve window " + windowId); return; }

                // CustomTaskPaneCollection.Add takes a UserControl; wrap the WPF ElementHost in one
                var wrapper = new PaneHostControl();
                var tabBar = new TabBarControl { DataContext = _vm };
                wrapper.ElementHost.Child = tabBar;

                var pane = _addIn.CustomTaskPanes.Add(wrapper, "CrabOffice", win);
                pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                pane.Height = 36;
                pane.Visible = true;

                _panes[windowId] = pane;
                ThisAddIn.SafeLog("pane created top-docked for win " + windowId);
            }
            catch (Exception ex)
            {
                ThisAddIn.SafeLog("EnsurePane FAILED for " + windowId + ": " + ex);
            }
        }

        /// <summary>Called by the adapter when a brand-new window appears (pane auto-provision).</summary>
        public void NotifyNewWindow(Guid windowId) { EnsurePane(windowId); }

        private Word.Window ResolveWindow(Guid windowId)
        {
            // windows are looked up through the adapter's live map via Globals; simplest: enumerate app windows by id tracked in adapter
            return ThisAddIn.Adapter != null ? ThisAddIn.Adapter.GetWindow(windowId) : null;
        }

        public void RemovePane(Guid windowId)
        {
            CustomTaskPane pane;
            if (!_panes.TryGetValue(windowId, out pane)) return;
            try { _addIn.CustomTaskPanes.Remove(pane); } catch { }
            _panes.Remove(windowId);
        }

        public void Dispose() { DisposeAll(); }

        public void DisposeAll()
        {
            foreach (var k in new List<Guid>(_panes.Keys)) RemovePane(k);
        }
    }
}
