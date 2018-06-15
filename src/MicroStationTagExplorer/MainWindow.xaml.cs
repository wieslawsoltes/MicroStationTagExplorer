using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Xml.Serialization;
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

        private ObservableCollection<File> GetFiles()
        {
            ObservableCollection<File> files = null;

            if (DataGridFiles.DataContext != null)
            {
                files = DataGridFiles.DataContext as ObservableCollection<File>;
            }
            else
            {
                files = new ObservableCollection<File>();
                DataGridFiles.DataContext = files;
            }

            return files;
        }

        private void SetFiles(ObservableCollection<File> files)
        {
            DataGridFiles.DataContext = files;
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
                ObservableCollection<File> files = GetFiles();

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

        private void AddPaths(string[] paths)
        {
            ObservableCollection<File> files = GetFiles();

            foreach (var path in paths)
            {
                System.IO.FileAttributes attributes = System.IO.File.GetAttributes(path);
                if (attributes.HasFlag(System.IO.FileAttributes.Directory))
                {
                    var fileNames = System.IO.Directory.EnumerateFiles(path, "*.*", System.IO.SearchOption.AllDirectories)
                                                       .Where(s => s.EndsWith(".dwg") || s.EndsWith(".dgn")); ;
                    foreach (var fileName in fileNames)
                    {
                        var file = new File()
                        {
                            Path = fileName
                        };
                        files.Add(file);
                    }
                }
                else
                {
                    if (path.EndsWith(".dwg") || path.EndsWith(".dgn"))
                    {
                        var file = new File()
                        {
                            Path = path
                        };
                        files.Add(file);
                    }
                }
            }
        }

        private void NewProject()
        {
            SetFiles(new ObservableCollection<File>());
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
                using (var reader = new System.IO.StreamReader(dlg.FileName))
                {
                    var serializer = new XmlSerializer(typeof(ObservableCollection<File>));
                    ObservableCollection<File> files = (ObservableCollection<File>)serializer.Deserialize(reader);
                    SetFiles(files);
                }
            }
        }

        private void SaveProject()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Supported Files (*.xml)|*.xml|All Files (*.*)|*.*",
                FileName = "tags"
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                ObservableCollection<File> files = GetFiles();

                using (var writer = new System.IO.StreamWriter(dlg.FileName))
                {
                    var serializer = new XmlSerializer(typeof(ObservableCollection<File>));
                    serializer.Serialize(writer, files);
                }
            }
        }

        private void Exit()
        {
            Close();
        }

        private void GetTags()
        {
            if (IsRunning == false)
            {
                IsRunning = true;

                DisableWindow();
                UpdateStatus("");

                TokenSource = new CancellationTokenSource();
                Token = TokenSource.Token;

                ObservableCollection<File> files = GetFiles();

                Task.Factory.StartNew(() =>
                {
                    try
                    {
                        Token.ThrowIfCancellationRequested();

                        int count = files.Count;
                        for (int i = 0; i < files.Count; i++)
                        {
                            var file = files[i];
                            if (Token.IsCancellationRequested)
                            {
                                Token.ThrowIfCancellationRequested();
                                return;
                            }

                            Dispatcher.Invoke(() =>
                            {
                                UpdateStatus("[" + i + "/" + count + "] " + System.IO.Path.GetFileName(file.Path));
                            });

                            using (var microstation = new MicrostationInterop(file.Path))
                            {
                                microstation.SetNormalActiveModel();
                                file.TagSets = microstation.GetTagSets();
                                file.Tags = microstation.GetTags();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                        Debug.WriteLine(ex.StackTrace);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        EnableWindow();
                        UpdateStatus("");
                    });

                    TokenSource.Dispose();
                    IsRunning = false;
                }, Token);
            }
            else
            {
                TokenSource.Cancel();
                TokenSource.Dispose();
                EnableWindow();
                UpdateStatus("");
            }
        }

        private void UpdateStatus(string message)
        {
            TextBoxStatus.Text = message;
        }

        private void EnableWindow()
        {
            FileMenu.IsEnabled = true;
            TagsImport.IsEnabled = true;
            TagsExport.IsEnabled = true;
            TagsGet.Header = "_Get";
        }

        private void DisableWindow()
        {
            FileMenu.IsEnabled = false;
            TagsImport.IsEnabled = false;
            TagsExport.IsEnabled = false;
            TagsGet.Header = "S_top";
        }

        private void ImportTags()
        {
        }

        private void ToValues(Tag[] tags, out object[,] values)
        {
            values = new object[tags.Length + 1, 6];
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
        }

        private void ExportTags()
        {
            if (IsRunning == false)
            {
                ObservableCollection<File> files = GetFiles();

                try
                {
                    object[,] values;
                    Tag[] tags = files.SelectMany(f => f.Tags).ToArray();
                    ToValues(tags, out values);
                    using (var excel = new Excelnterop())
                    {
                        excel.ExportTags(values, tags.Length + 1, 6);
                    }
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
                if (e.Key == Key.L)
                {
                    AddFiles();
                }
                else if(e.Key == Key.N)
                {
                    NewProject();
                }
                else if (e.Key == Key.O)
                {
                    OpenProject();
                }
                else if (e.Key == Key.S)
                {
                    SaveProject();
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

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop);
                if (data != null && data is string[])
                {
                    AddPaths(data as string[]);
                }
            }
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.All;
            }
        }

        private void FileAddFiles_Click(object sender, RoutedEventArgs e)
        {
            AddFiles();
        }

        private void FileNewProject_Click(object sender, RoutedEventArgs e)
        {
            NewProject();
        }

        private void FileOpenProject_Click(object sender, RoutedEventArgs e)
        {
            OpenProject();
        }

        private void FileSaveProject_Click(object sender, RoutedEventArgs e)
        {
            SaveProject();
        }

        private void FileExit_Click(object sender, RoutedEventArgs e)
        {
            Exit();
        }

        private void TagsGet_Click(object sender, RoutedEventArgs e)
        {
            GetTags();
        }

        private void TagsImport_Click(object sender, RoutedEventArgs e)
        {
            ImportTags();
        }

        private void TagsExport_Click(object sender, RoutedEventArgs e)
        {
            ExportTags();
        }
    }
}
