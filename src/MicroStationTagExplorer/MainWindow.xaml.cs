using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using MicroStationTagExplorer.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void FileAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "All Files (*.*)|*.*",
                Multiselect = true
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                IList<DgnFile> dgnFiles = null;
                if (DataGridFiles.DataContext != null)
                {
                    dgnFiles = DataGridFiles.DataContext as IList<DgnFile>;
                }
                else
                {
                    dgnFiles = new List<DgnFile>();
                }

                foreach (var fileName in dlg.FileNames)
                {
                    var dgnFile = new DgnFile()
                    {
                        Path = fileName
                    };
                    dgnFiles.Add(dgnFile);
                }

                DataGridFiles.DataContext = dgnFiles;
            }
        }

        private void FileGetTags_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridFiles.DataContext != null)
            {
                int selectedIndex = DataGridFiles.SelectedIndex;
                IList<DgnFile> dgnFiles = DataGridFiles.DataContext as IList<DgnFile>;
                if (dgnFiles != null)
                {
                    MainMenu.IsEnabled = false;
                    TextBlockStatus.Text = "";
                    Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            foreach (var dgnFile in dgnFiles)
                            {
                                Dispatcher.Invoke(() => TextBlockStatus.Text = dgnFile.Path);
                                Microstation.GetTagData(dgnFile);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            Debug.WriteLine(ex.StackTrace);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            TextBlockStatus.Text = "";
                            MainMenu.IsEnabled = true;
                        });
                    });
                }
            }
        }

        private void FileImportTags_Click(object sender, RoutedEventArgs e)
        {

        }

        private void FileExportTags_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application app = Utilities.CreateObject<Excel.Application>("Excel.Application");
            app.Visible = true;

            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets.Add();
        }
    }
}
