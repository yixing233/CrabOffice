using System;
using System.IO;
using CrabOffice.Core;
using CrabOffice.UI;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace CrabOffice.WordHost
{
    public partial class ThisAddIn
    {
        internal static DocumentRegistry Registry;
        internal static TabService Tabs;
        internal static TabBarViewModel ViewModel;
        internal static WordHostAdapter Adapter;
        internal static TabPanePresenter Presenter;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                SafeLog("startup begin");
                Registry = new DocumentRegistry();
                Adapter = new WordHostAdapter(this);
                Adapter.HostEvent += ev => Registry.ApplyHostEvent(ev);
                Tabs = new TabService(Registry, Adapter);
                ViewModel = new TabBarViewModel(Registry, Tabs, id =>
                {
                    var d = Registry.GetDoc(id);
                    return d != null ? d.DisplayName : "?";
                });
                Presenter = new TabPanePresenter(this, ViewModel);

                // initial snapshot: documents already open when add-in loads (A02)
                var snap = Adapter.EnumerateSnapshot();
                Registry.Reconcile(snap);
                foreach (var win in snap.Windows) Presenter.EnsurePane(win.WindowId);
                SafeLog("startup ok, windows=" + snap.Windows.Count);
            }
            catch (Exception ex)
            {
                SafeLog("startup FAILED: " + ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Presenter?.DisposeAll();
                Registry?.ApplyHostEvent(new DomainEvent(EventType.HostShutdown, Guid.Empty));
                SafeLog("shutdown ok");
            }
            catch (Exception ex) { SafeLog("shutdown FAILED: " + ex); }
        }

        internal static void SafeLog(string msg)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CrabOffice");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "wordhost.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg + Environment.NewLine);
            }
            catch { /* never break Word because of logging */ }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
