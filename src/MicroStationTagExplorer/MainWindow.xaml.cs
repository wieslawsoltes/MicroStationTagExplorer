using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

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
                    files = new ObservableCollection<File>();
                    DataGridFiles.DataContext = files;
                }

                foreach (var fileName in dlg.FileNames)
                {
                    var file = new File()
                    {
                        Path = fileName
                    };
                    files.Add(file);
                }
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
                    try
                    {
                        Token.ThrowIfCancellationRequested();

                        foreach (var file in files)
                        {
                            if (Token.IsCancellationRequested)
                            {
                                Token.ThrowIfCancellationRequested();
                                return;
                            }

                            Dispatcher.Invoke(() =>
                            {
                                TextBoxStatus.Text = System.IO.Path.GetFileName(file.Path);
                            });

                            MicrostationInterop.GetTagData(file);
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
            if (DataGridFiles.DataContext == null)
                return;

            if (IsRunning == false)
            {
                int selectedIndex = DataGridFiles.SelectedIndex;
                IList<File> files = DataGridFiles.DataContext as IList<File>;
                if (files == null)
                    return;

                try
                {
                    var tags = files.SelectMany(f => f.Tags).ToArray();
                    var values = new object[tags.Length + 1, 6];

                    values[0, 0] = "TagSetName";
                    values[0, 1] = "TagDefinitionName";
                    values[0, 2] = "Value";
                    values[0, 3] = "ID";
                    values[0, 4] = "HostID";
                    values[0, 5] = "Path";

                    for (int i = 0; i < tags.Length; i++)
                    {
                        values[i + 1, 0] = tags[i].TagSetName;
                        values[i + 1, 1] = tags[i].TagDefinitionName;
                        values[i + 1, 2] = tags[i].Value.ToString();
                        values[i + 1, 3] = tags[i].ID.ToString();
                        values[i + 1, 4] = tags[i].HostID.ToString();
                        values[i + 1, 5] = tags[i].Path;
                    }

                    //Excelnterop.ExportTags(values, tags.Length + 1, 6);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Debug.WriteLine(ex.StackTrace);
                }
            }
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
