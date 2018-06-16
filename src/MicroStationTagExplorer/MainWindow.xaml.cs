using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace MicroStationTagExplorer
{
    public partial class MainWindow : Window
    {
        private static StringComparison _comparisonType = StringComparison.OrdinalIgnoreCase;
        private static string _xmlExt = ".xml";
        private static string _dgnExt = ".dgn";
        private static string _dwgExt = ".dwg";
        private volatile bool IsRunning = false;
        private CancellationTokenSource TokenSource;
        private CancellationToken Token;
        private Project _project;

        public MainWindow()
        {
            InitializeComponent();
            NewProject();
        }

        private void UpdateProject(Project project)
        {
            foreach (var file in project.Files)
            {
                ValidateFile(file);
            }
        }

        private void ValidateFile(File file)
        {
            foreach (var tagSet in file.TagSets)
            {
                tagSet.File = file;
            }

            foreach (var tag in file.Tags)
            {
                tag.File = file;
            }

            var elementsByHostID = file.Tags.GroupBy(t => t.HostID)
                                             .Select(g => new Element<Int64>()
                                             {
                                                 Key = g.Key,
                                                 Tags = new ObservableCollection<Tag>(g)
                                             });

            var elementsByTagSet = file.Tags.GroupBy(t => t.HostID)
                                             .Select(e => e.GroupBy(t => t.TagSetName))
                                             .SelectMany(e => e)
                                             .Select(g => new Element<string>()
                                             {
                                                 Key = g.Key,
                                                 Tags = new ObservableCollection<Tag>(g)
                                             });

            file.ElementsByHostID = new ObservableCollection<Element<Int64>>(elementsByHostID);
            file.ElementsByTagSet = new ObservableCollection<Element<string>>(elementsByTagSet);
            file.Errors = new ObservableCollection<Error>(ValidateTags(file));
            file.HasErrors = file.Errors.Count > 0;
        }

        private IEnumerable<Error> ValidateTags(File file)
        {
            foreach (var element in file.ElementsByTagSet)
            {
                var tagSet = file.TagSets.FirstOrDefault(ts => ts.Name == element.Key);
                var errors = ValidateTags(element, tagSet);
                element.Errors = new ObservableCollection<Error>(errors);
                element.HasErrors = element.Errors.Count > 0;
                foreach (var error in element.Errors)
                {
                    error.File = file;
                    yield return error;
                }
            }
        }

        private IEnumerable<Error> ValidateTags(Element element, TagSet tagSet)
        {
            var first = element.Tags.First();
            bool isSameTagSet = element.Tags.All(e => e.TagSetName == first.TagSetName);

            if (element.Tags.Count() != tagSet.TagDefinitions.Count)
            {
                yield return new Error()
                {
                    Message = "Missing Tags: " + tagSet.Name,
                    Element = element,
                    TagSet = tagSet
                };
            }

            // check for missing tags from tag set

            foreach (var tagDef in tagSet.TagDefinitions)
            {
                bool isValidTag = false;

                foreach (var tag in element.Tags)
                {
                    if (tag.TagDefinitionName == tagDef.Name)
                    {
                        isValidTag = true;
                    }
                }
                if (isValidTag == false)
                {
                    yield return new Error()
                    {
                        Message = "Missing Tag: " + tagDef.Name + " [" + tagSet.Name + "]",
                        Element = element,
                        TagSet = tagSet
                    };
                }
            }

            // check for invalid tags in element

            foreach (var tag in element.Tags)
            {
                bool isValidElement = false;

                foreach (var tagDef in tagSet.TagDefinitions)
                {
                    if (tagDef.Name == tag.TagDefinitionName)
                    {
                        isValidElement = true;
                    }
                }
                if (isValidElement == false)
                {
                    yield return new Error()
                    {
                        Message = "Invalid Tag: " + tag.TagDefinitionName + " [" + tagSet.Name + "]",
                        Element = element,
                        TagSet = tagSet
                    };
                }
            }
        }

        private bool IsProject(string path)
        {
            return path.EndsWith(_xmlExt, _comparisonType);
        }

        private bool IsSupportedFile(string path)
        {
            return path.EndsWith(_dwgExt, _comparisonType) || path.EndsWith(_dgnExt, _comparisonType);
        }

        private void AddFile(string path)
        {
            if (IsSupportedFile(path))
            {
                var file = new File()
                {
                    Name = System.IO.Path.GetFileName(path),
                    Path = path
                };
                _project.Files.Add(file);
            }
        }

        private void AddFiles(string[] fileNames)
        {
            foreach (var fileName in fileNames)
            {
                AddFile(fileName);
            }
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
                AddFiles(dlg.FileNames);
            }
        }

        private void AddPaths(string[] paths)
        {
            foreach (var path in paths)
            {
                System.IO.FileAttributes attributes = System.IO.File.GetAttributes(path);
                if (attributes.HasFlag(System.IO.FileAttributes.Directory))
                {
                    var fileNames = System.IO.Directory.EnumerateFiles(path, "*.*", System.IO.SearchOption.AllDirectories)
                                                       .Where(s => IsSupportedFile(s));
                    foreach (var fileName in fileNames)
                    {
                        AddFile(fileName);
                    }
                }
                else
                {
                    AddFile(path);
                }
            }
        }

        private void NewProject()
        {
            _project = new Project()
            {
                Name = "project",
                Path = "projet" + _xmlExt,
                Files = new ObservableCollection<File>()
            };

            UpdateProject(_project);

            this.DataContext = _project;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void SerializeXmlSerializer(Project project, string fileName)
        {
            using (var writer = new System.IO.StreamWriter(fileName))
            {
                var serializer = new XmlSerializer(typeof(Project));
                serializer.Serialize(writer, project);
            }
        }

        private Project DeserializeXmlSerializer(string fileName)
        {
            Project project = null;
            using (var reader = new System.IO.StreamReader(fileName))
            {
                var serializer = new XmlSerializer(typeof(Project));
                project = (Project)serializer.Deserialize(reader);
            }
            return project;
        }

        private void SerializeDataContract(Project project, string fileName)
        {
            using (var stream = new System.IO.StreamWriter(fileName))
            {
                var settings = new XmlWriterSettings()
                {
                    Indent = true,
                    IndentChars = "    ",
                    Encoding = Encoding.UTF8,
                };
                using (var writer = XmlWriter.Create(stream, settings))
                {
                    var serializer = new DataContractSerializer(typeof(Project), null, int.MaxValue, false, true, null);
                    serializer.WriteObject(writer, project);
                }
            }
        }

        private Project DeserializeDataContract(string fileName)
        {
            Project project = null;
            using (var stream = new System.IO.StreamReader(fileName))
            {
                var settings = new XmlReaderSettings();
                using (var reader = XmlReader.Create(stream, settings))
                {
                    var serializer = new DataContractSerializer(typeof(Project), null, int.MaxValue, false, true, null);
                    project = (Project)serializer.ReadObject(reader);
                }
            }
            return project;
        }

        private void OpenProject(string fileName)
        {
            Project project = DeserializeXmlSerializer(fileName);
            if (project != null)
            {
                _project = project;
                _project.Path = fileName;

                foreach (var file in _project.Files)
                {
                    file.Name = System.IO.Path.GetFileName(file.Path);
                }

                UpdateProject(_project);

                this.DataContext = _project;
            }
        }

        private void SaveProject(string fileName)
        {
            SerializeXmlSerializer(_project, fileName);
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
                OpenProject(dlg.FileName);
            }
        }

        private void SaveProject()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Supported Files (*.xml)|*.xml|All Files (*.*)|*.*",
                FileName = _project.Path
            };
            var result = dlg.ShowDialog(this);
            if (result == true)
            {
                SaveProject(dlg.FileName);
            }
        }

        private void Exit()
        {
            Close();
        }

        private void GetTags(Func<File, int, int, bool> updateStatus, IEnumerable<File> files)
        {
            int count = files.Count();
            int index = 0;

            foreach (var file in _project.Files)
            {
                index++;
                if (updateStatus(file, index, count) == false)
                    return;

                using (var microstation = new MicrostationInterop(file.Path))
                {
                    microstation.SetNormalActiveModel();
                    file.TagSets = microstation.GetTagSets();
                    file.Tags = microstation.GetTags();
                    ValidateFile(file);
                }
            }
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

                Func<File, int, int, bool> updateStatus = (file, current, total) =>
                 {
                     if (Token.IsCancellationRequested)
                     {
                         Token.ThrowIfCancellationRequested();
                         return false;
                     }

                     Dispatcher.Invoke(() =>
                     {
                         UpdateStatus("[" + current + "/" + total + "] " + System.IO.Path.GetFileName(file.Path));
                     });

                     return true;
                 };

                Task.Factory.StartNew(() =>
                {
                    try
                    {
                        Token.ThrowIfCancellationRequested();
                        GetTags(updateStatus, _project.Files);
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
            // TODO: Import tags from excel file.
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
                values[i + 1, 5] = tags[i].File.Path;
            }
        }

        private void ExportTags()
        {
            try
            {
                object[,] values;
                Tag[] tags = _project.Files.SelectMany(f => f.Tags).ToArray();
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

        private void DropFiles(string[] paths)
        {
            if (paths.Length == 1 && IsProject(paths[0]))
            {
                OpenProject(paths[0]);
            }
            else
            {
                AddPaths(paths);
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
                    if (IsRunning == false)
                    {
                        AddFiles();
                    }
                }
                else if (e.Key == Key.N)
                {
                    if (IsRunning == false)
                    {
                        NewProject();
                    }
                }
                else if (e.Key == Key.O)
                {
                    if (IsRunning == false)
                    {
                        OpenProject();
                    }
                }
                else if (e.Key == Key.S)
                {
                    if (IsRunning == false)
                    {
                        SaveProject();
                    }
                }
                else if (e.Key == Key.I)
                {
                    if (IsRunning == false)
                    {
                        ImportTags();
                    }
                }
                else if (e.Key == Key.E)
                {
                    if (IsRunning == false)
                    {
                        ExportTags();
                    }
                }
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (IsRunning == false)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var data = e.Data.GetData(DataFormats.FileDrop);
                    if (data != null && data is string[])
                    {
                        var paths = data as string[];
                        DropFiles(paths);
                    }
                }
            }
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (IsRunning == false)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effects = DragDropEffects.All;
                }
            }
        }

        private void FileAddFiles_Click(object sender, RoutedEventArgs e)
        {
            if (IsRunning == false)
            {
                AddFiles();
            }
        }

        private void FileNewProject_Click(object sender, RoutedEventArgs e)
        {
            if (IsRunning == false)
            {
                NewProject();
            }
        }

        private void FileOpenProject_Click(object sender, RoutedEventArgs e)
        {
            if (IsRunning == false)
            {
                OpenProject();
            }
        }

        private void FileSaveProject_Click(object sender, RoutedEventArgs e)
        {
            if (IsRunning == false)
            {
                SaveProject();
            }
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
            if (IsRunning == false)
            {
                ImportTags();
            }
        }

        private void TagsExport_Click(object sender, RoutedEventArgs e)
        {
            if (IsRunning == false)
            {
                ExportTags();
            }
        }
    }
}
