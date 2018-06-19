using System.Windows.Controls;

namespace MicroStationTagExplorer.Views
{
    public partial class SheetsControl : UserControl
    {
        public SheetsControl()
        {
            InitializeComponent();
        }

        private void SetIsCheckedValue(bool value)
        {
            foreach (var item in DataGridSheets.ItemsSource)
            {
                if (item is Sheet)
                {
                    Sheet sheet = item as Sheet;
                    sheet.IsExported = value;
                }

                var content = DataGridSheets.Columns[0].GetCellContent(item);
                if (content != null && content is CheckBox)
                {
                    ((CheckBox)content).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
                }
            }
        }

        private void ButtonSelectNone_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            SetIsCheckedValue(false);
        }

        private void ButtonSelectAll_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            SetIsCheckedValue(true);
        }
    }
}
