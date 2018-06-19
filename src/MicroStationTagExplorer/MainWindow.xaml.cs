using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace MicroStationTagExplorer
{
    public partial class MainWindow : Window
    {
        public TagExplorer Explorer { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Explorer = new TagExplorer();
            NewProject();
            SettingsWorkers.Text = Explorer.WorkersNum.ToString();
        }

        private void UpdateStatus(string message)
        {
            TextBoxStatus.Text = message;
        }

        private void AddFiles()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Supported Files (*.dwg;*dgn)|*.dwg;*dgn|All Files (*.*)|*.*",
                Multiselect = true
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                Explorer.AddFiles(dlg.FileNames);
            }
        }

        private void NewProject()
        {
            Explorer.NewProject();
            DataContext = Explorer.Project;
        }

        private void OpenProject()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Supported Files (*.xml)|*.xml|All Files (*.*)|*.*",
                Multiselect = false
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                Explorer.OpenProject(dlg.FileName);
                DataContext = Explorer.Project;
            }
        }

        private void SaveProject()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Supported Files (*.xml)|*.xml|All Files (*.*)|*.*",
                FileName = Explorer.Project.Path
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                Explorer.SaveProject(dlg.FileName);
            }
        }

        private async Task GetTags()
        {
            if (Explorer.IsRunning == false)
            {
                if (Explorer.Project.Files.Count <= 0)
                    return;

                Explorer.WorkersNum = int.Parse(SettingsWorkers.Text);
                if (Explorer.WorkersNum <= 0)
                {
                    Explorer.WorkersNum = 1;
                }

                Explorer.ActiveWorkers.Clear();
                foreach (var worker in Explorer.Workers)
                {
                    if (worker.IsEnabled)
                    {
                        Explorer.ActiveWorkers.Add(worker);
                    }
                }

                if (Explorer.ActiveWorkers.Count > 0)
                {
                    Explorer.WorkersNum = Explorer.ActiveWorkers.Count;
                }

                Explorer.IsRunning = true;

                DisableWindow();
                UpdateStatus("");

                Explorer.TokenSources = new List<CancellationTokenSource>();
                Explorer.Tokens = new List<CancellationToken>();

                for (int i = 0; i < Explorer.WorkersNum; i++)
                {
                    var tokenSource = new CancellationTokenSource();
                    Explorer.TokenSources.Add(tokenSource);
                    Explorer.Tokens.Add(tokenSource.Token);
                }

                int nPreviousCurrent = -1;
                int nCurrent = 0;

                Func<File, int, bool> updateStatus = (file, id) =>
                {
                    if (Explorer.Tokens[id].IsCancellationRequested)
                    {
                        return false;
                    }

                    nCurrent++;
                    if (nPreviousCurrent < nCurrent)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            UpdateStatus("[" + nCurrent + "/" + Explorer.Project.Files.Count + "] ");
                        });
                        nPreviousCurrent = nCurrent;
                    }

                    return true;
                };

                int count = (int)Math.Ceiling(Explorer.Project.Files.Count / (double)Explorer.WorkersNum);
                var partitions = Explorer.Split(Explorer.Project.Files, count);
                var tasks = new List<Task>();

                for (int i = 0; i < Explorer.WorkersNum; i++)
                {
                    int id = i;
                    var task = Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            Explorer.Tokens[id].ThrowIfCancellationRequested();
                            if (Explorer.ActiveWorkers.Count > 0)
                            {
                                Explorer.GetTags(updateStatus, partitions[id], id, false, Explorer.ActiveWorkers[id]);
                            }
                            else
                            {
                                Explorer.GetTags(updateStatus, partitions[id], id, Explorer.WorkersNum > 1, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            Debug.WriteLine(ex.StackTrace);
                        }
                    }, Explorer.Tokens[id]);
                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);

                Dispatcher.Invoke(() =>
                {
                    EnableWindow();
                    UpdateStatus("");
                });

                for (int i = 0; i < Explorer.WorkersNum; i++)
                {
                    Explorer.TokenSources[i].Dispose();
                }

                Explorer.IsRunning = false;
            }
            else
            {
                for (int i = 0; i < Explorer.WorkersNum; i++)
                {
                    Explorer.TokenSources[i].Cancel();
                    Explorer.TokenSources[i].Dispose();
                }
                EnableWindow();
                UpdateStatus("");
            }
        }

        private void ImportTags()
        {
            try
            {
                Explorer.ImportTags();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void ExportTags()
        {
            try
            {
                Explorer.ExportTags();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void DropFiles(string[] paths)
        {
            if (paths.Length == 1 && Explorer.IsProject(paths[0]))
            {
                Explorer.OpenProject(paths[0]);
                DataContext = Explorer.Project;
            }
            else
            {
                Explorer.AddPaths(paths);
            }
        }

        private void GetWorkers()
        {
            try
            {
                Explorer.Workers = new List<Worker>();
                Explorer.GetWorkers();
                WorkersView.DataGridWorkers.DataContext = Explorer.Workers;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void Exit()
        {
            Close();
        }

        private void EnableWindow()
        {
            FileMenu.IsEnabled = true;
            SettingsMenu.IsEnabled = true;
            TagsImport.IsEnabled = true;
            TagsExport.IsEnabled = true;
            TagsGet.Header = "_Get";
        }

        private void DisableWindow()
        {
            FileMenu.IsEnabled = false;
            SettingsMenu.IsEnabled = false;
            TagsImport.IsEnabled = false;
            TagsExport.IsEnabled = false;
            TagsGet.Header = "S_top";
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            bool bNone = Keyboard.Modifiers == ModifierKeys.None;
            bool bControl = Keyboard.Modifiers == ModifierKeys.Control;
            if (bNone)
            {
                if (e.Key == Key.F5)
                {
                    await GetTags();
                }
            }
            else if (bControl)
            {
                if (e.Key == Key.L)
                {
                    if (Explorer.IsRunning == false)
                    {
                        AddFiles();
                    }
                }
                else if (e.Key == Key.N)
                {
                    if (Explorer.IsRunning == false)
                    {
                        Explorer.NewProject();
                        DataContext = Explorer.Project;
                    }
                }
                else if (e.Key == Key.O)
                {
                    if (Explorer.IsRunning == false)
                    {
                        OpenProject();
                    }
                }
                else if (e.Key == Key.S)
                {
                    if (Explorer.IsRunning == false)
                    {
                        SaveProject();
                    }
                }
                else if (e.Key == Key.I)
                {
                    if (Explorer.IsRunning == false)
                    {
                        ImportTags();
                    }
                }
                else if (e.Key == Key.E)
                {
                    if (Explorer.IsRunning == false)
                    {
                        ExportTags();
                    }
                }
                else if (e.Key == Key.W)
                {
                    if (Explorer.IsRunning == false)
                    {
                        GetWorkers();
                    }
                }
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var data = e.Data.GetData(DataFormats.FileDrop);
                    if (data != null && data is string[])
                    {
                        DropFiles(data as string[]);
                    }
                }
            }
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effects = DragDropEffects.All;
                }
            }
        }

        private void FileAddFiles_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                AddFiles();
            }
        }

        private void FileNewProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                NewProject();
            }
        }

        private void FileOpenProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                OpenProject();
            }
        }

        private void FileSaveProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                SaveProject();
            }
        }

        private void FileExit_Click(object sender, RoutedEventArgs e)
        {
            Exit();
        }

        private async void TagsGet_Click(object sender, RoutedEventArgs e)
        {
            await GetTags();
        }

        private void TagsImport_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ImportTags();
            }
        }

        private void TagsExport_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ExportTags();
            }
        }

        private void SettingsGetWorkers_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                GetWorkers();
            }
        }
    }
}
