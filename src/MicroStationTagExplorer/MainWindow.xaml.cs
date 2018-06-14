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
                IList<File> files = null;
                if (DataGridFiles.DataContext != null)
                {
                    files = DataGridFiles.DataContext as IList<File>;
                }
                else
                {
                    files = new List<File>();
                }

                foreach (var fileName in dlg.FileNames)
                {
                    var file = new File()
                    {
                        Path = fileName
                    };
                    files.Add(file);
                }

                DataGridFiles.DataContext = files;
            }
        }

        private void FileGetTags_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridFiles.DataContext != null)
            {
                int selectedIndex = DataGridFiles.SelectedIndex;
                IList<File> files = DataGridFiles.DataContext as IList<File>;
                if (files != null)
                {
                    MainMenu.IsEnabled = false;
                    TextBoxStatus.Text = "";
                    Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            foreach (var dgnFile in files)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    TextBoxStatus.Text = System.IO.Path.GetFileName(dgnFile.Path);
                                });
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
                            TextBoxStatus.Text = "";
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
