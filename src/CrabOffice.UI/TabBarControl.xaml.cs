using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CrabOffice.UI
{
    public partial class TabBarControl : UserControl
    {
        public TabBarControl()
        {
            InitializeComponent();
            Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("pack://application:,,,/CrabOffice.UI;component/TabBarTheme.xaml") });
            DataContextChanged += (s, e) => { };
        }

        private TabBarViewModel Vm { get { return DataContext as TabBarViewModel; } }

        private void OnTabMouseDown(object sender, MouseButtonEventArgs e)
        {
            var vm = (sender as FrameworkElement)?.DataContext as TabItemViewModel;
            if (vm != null && Vm != null) Vm.Activate(vm.TabId);
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            var vm = ((sender as Button)?.DataContext) as TabItemViewModel;
            if (vm != null && Vm != null) Vm.Close(vm.TabId);
        }
    }
}
