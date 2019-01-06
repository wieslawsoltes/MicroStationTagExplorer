using System.Windows.Controls;
using MicroStationTagExplorer.Core.Model;

namespace MicroStationTagExplorer.Views
{
    public partial class SheetsControl : UserControl
    {
        public SheetsControl()
        {
            InitializeComponent();
        }

        private void SetItemIsCheckedValue(object item, bool value)
        {
            if (item is Sheet)
            {
                Sheet sheet = item as Sheet;
                sheet.IsExported = value;

                var content = DataGridSheets.Columns[0].GetCellContent(item);
                if (content != null && content is CheckBox)
                {
                    ((CheckBox)content).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
                }
            }
        }

        private void ButtonSelectNone_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            foreach (var item in DataGridSheets.ItemsSource)
            {
                SetItemIsCheckedValue(item, false);
            }
        }

        private void ButtonSelectAll_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            foreach (var item in DataGridSheets.ItemsSource)
            {
                SetItemIsCheckedValue(item, true);
            }
        }

        private void ButtonDeselect_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            foreach (var item in DataGridSheets.SelectedItems)
            {
                SetItemIsCheckedValue(item, false);
            }
        }

        private void ButtonSelect_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            foreach (var item in DataGridSheets.SelectedItems)
            {
                SetItemIsCheckedValue(item, true);
            }
        }
    }
}
