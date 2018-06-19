﻿using System;
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
            NewProjectImpl();
            SettingsWorkers.Text = Explorer.WorkersNum.ToString();
        }

        private void UpdateStatus(string message)
        {
            TextBoxStatus.Text = message;
        }

        private void AddFilesImpl()
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

        private void NewProjectImpl()
        {
            Explorer.NewProject();
            DataContext = Explorer.Project;
        }

        private void OpenProjectImpl()
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

        private void SaveProjectImpl()
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

        private async Task GetDataImpl()
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
                                Explorer.GetData(updateStatus, partitions[id], id, false, Explorer.ActiveWorkers[id]);
                            }
                            else
                            {
                                Explorer.GetData(updateStatus, partitions[id], id, Explorer.WorkersNum > 1, null);
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

                Explorer.ValidateProject(Explorer.Project);

                DataContext = null;
                DataContext = Explorer.Project;

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

        private void ResetDataImpl()
        {
            var result = MessageBox.Show("Reset data?", "Warning", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                Explorer.ResetData();
                DataContext = null;
                DataContext = Explorer.Project;
            }
        }

        private void ImportTagsImpl()
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

        private void ImportElementsImpl()
        {
            try
            {
                Explorer.ImportElements();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void ImportTextsImpl()
        {
            try
            {
                Explorer.ImportTexts();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void ExportTagsImpl()
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

        private void ExportElementsImpl()
        {
            try
            {
                Explorer.ExportElements();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void ExportTextsImpl()
        {
            try
            {
                Explorer.ExportTexts();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private void DropFilesImpl(string[] paths)
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

        private void GetWorkersImpl()
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

        private void ExitImpl()
        {
            Close();
        }

        private void EnableWindow()
        {
            FileMenu.IsEnabled = true;
            DataGet.Header = "_Get";
            DataReset.IsEnabled = true;
            ImportMenu.IsEnabled = true;
            ExportMenu.IsEnabled = true;
            SettingsMenu.IsEnabled = true;
        }

        private void DisableWindow()
        {
            FileMenu.IsEnabled = false;
            DataGet.Header = "S_top";
            DataReset.IsEnabled = false;
            ImportMenu.IsEnabled = false;
            ExportMenu.IsEnabled = false;
            SettingsMenu.IsEnabled = false;
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            bool bNone = Keyboard.Modifiers == ModifierKeys.None;
            bool bControl = Keyboard.Modifiers == ModifierKeys.Control;
            if (bNone)
            {
                if (e.Key == Key.F5)
                {
                    await GetDataImpl();
                }
            }
            else if (bControl)
            {
                if (e.Key == Key.L)
                {
                    if (Explorer.IsRunning == false)
                    {
                        AddFilesImpl();
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
                        OpenProjectImpl();
                    }
                }
                else if (e.Key == Key.S)
                {
                    if (Explorer.IsRunning == false)
                    {
                        SaveProjectImpl();
                    }
                }
                else if (e.Key == Key.I)
                {
                    if (Explorer.IsRunning == false)
                    {
                        ImportTagsImpl();
                    }
                }
                else if (e.Key == Key.E)
                {
                    if (Explorer.IsRunning == false)
                    {
                        ExportTagsImpl();
                    }
                }
                else if (e.Key == Key.W)
                {
                    if (Explorer.IsRunning == false)
                    {
                        GetWorkersImpl();
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
                        DropFilesImpl(data as string[]);
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
                AddFilesImpl();
            }
        }

        private void FileNewProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                NewProjectImpl();
            }
        }

        private void FileOpenProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                OpenProjectImpl();
            }
        }

        private void FileSaveProject_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                SaveProjectImpl();
            }
        }

        private void FileExit_Click(object sender, RoutedEventArgs e)
        {
            ExitImpl();
        }

        private async void DataGet_Click(object sender, RoutedEventArgs e)
        {
            await GetDataImpl();
        }

        private void DataReset_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ResetDataImpl();
            }
        }
        private void ImportTags_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ImportTagsImpl();
            }
        }

        private void ImportElements_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ImportElementsImpl();
            }
        }

        private void ImportTexts_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ImportTextsImpl();
            }
        }

        private void ExportTags_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ExportTagsImpl();
            }
        }

        private void ExportElements_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ExportElementsImpl();
            }
        }

        private void ExportTexts_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                ExportTextsImpl();
            }
        }

        private void SettingsGetWorkers_Click(object sender, RoutedEventArgs e)
        {
            if (Explorer.IsRunning == false)
            {
                GetWorkersImpl();
            }
        }
    }
}
