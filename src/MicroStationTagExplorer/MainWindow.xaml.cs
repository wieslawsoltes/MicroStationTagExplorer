using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using MicroStationTagExplorer.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    public partial class MainWindow : Window
    {
        private volatile bool IsRunning = false;
        private CancellationTokenSource TokenSource;
        private CancellationToken Token;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddFiles()
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

                DataGridFiles.DataContext = null;
                DataGridFiles.DataContext = files;
            }
        }

        private void GetTags()
        {
            if (DataGridFiles.DataContext == null)
                return;

            if (IsRunning == false)
            {
                int selectedIndex = DataGridFiles.SelectedIndex;
                IList<File> files = DataGridFiles.DataContext as IList<File>;
                if (files == null)
                    return;

                IsRunning = true;
                FileGetTags.Header = "S_top";
                TextBoxStatus.Text = "";
                TokenSource = new CancellationTokenSource();
                Token = TokenSource.Token;

                Task.Factory.StartNew(() =>
                {
                    Token.ThrowIfCancellationRequested();
                    try
                    {
                        foreach (var dgnFile in files)
                        {
                            if (Token.IsCancellationRequested)
                            {
                                Token.ThrowIfCancellationRequested();
                                return;
                            }

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
                        FileGetTags.Header = "_Get Tags";
                        TextBoxStatus.Text = "";
                    });

                    TokenSource.Dispose();
                    IsRunning = false;
                }, Token);
            }
            else
            {
                TokenSource.Cancel();
                TokenSource.Dispose();
                FileGetTags.Header = "_Get Tags";
                TextBoxStatus.Text = "";
            }
        }

        private void ImportTags()
        {
        }

        private void ExportTags()
        {
            Excel.Application app = Utilities.CreateObject<Excel.Application>("Excel.Application");
            app.Visible = true;

            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets.Add();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            bool bNone = Keyboard.Modifiers == ModifierKeys.None;
            bool bControl = Keyboard.Modifiers == ModifierKeys.Control;
            if (bNone)
            {
                if (e.Key == Key.F5)
                {
                    GetTags();
                }
            }
            else if (bControl)
            {
                if (e.Key == Key.O)
                {
                    AddFiles();
                }
                else if (e.Key == Key.I)
                {
                    ImportTags();
                }
                else if (e.Key == Key.E)
                {
                    ExportTags();
                }
            }
        }

        private void FileAddFiles_Click(object sender, RoutedEventArgs e)
        {
            AddFiles();
        }

        private void FileGetTags_Click(object sender, RoutedEventArgs e)
        {
            GetTags();
        }

        private void FileImportTags_Click(object sender, RoutedEventArgs e)
        {
            ImportTags();
        }

        private void FileExportTags_Click(object sender, RoutedEventArgs e)
        {
            ExportTags();
        }
    }
}
