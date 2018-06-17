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
        private int Workers = 1;
        private List<CancellationTokenSource> TokenSources;
        private List<CancellationToken> Tokens;
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
                                                 File = file,
                                                 Tags = new ObservableCollection<Tag>(g)
                                             });

            var elementsByTagSet = file.Tags.GroupBy(t => t.HostID)
                                             .Select(e => e.GroupBy(t => t.TagSetName))
                                             .SelectMany(e => e)
                                             .Select(g => new Element<string>()
                                             {
                                                 Key = g.Key,
                                                 File = file,
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

        public IList<IList<T>> Split<T>(IList<T> list, int count)
        {
            var chunks = new List<IList<T>>();
            var chunkCount = list.Count() / count;

            if (list.Count % count > 0)
            {
                chunkCount++;
            }

            for (var i = 0; i < chunkCount; i++)
            {
                chunks.Add(list.Skip(i * count).Take(count).ToList());
            }

            return chunks;
        }

        private void GetTags(Func<File, int, bool> updateStatus, IEnumerable<File> files, int id, bool bConnect)
        {
            Debug.WriteLine("GetTags: " + id);

            using (var microstation = new MicrostationInterop())
            {
                if (bConnect == true)
                    microstation.ConnectApplication();
                else
                    microstation.CreateApplication();

                foreach (var file in files)
                {
                    if (updateStatus(file, id) == false)
                        break;

                    microstation.Open(file.Path);
                    microstation.SetNormalActiveModel();
                    file.TagSets = microstation.GetTagSets();
                    file.Tags = microstation.GetTags();
                    microstation.Close();
                    ValidateFile(file);
                }

                if (bConnect == true)
                {
                    microstation.Quit();
                }
            }
        }

        private async Task GetTags()
        {
            if (IsRunning == false)
            {
                Workers = int.Parse(SettingsWorkers.Text);
                if (Workers <= 0)
                {
                    Workers = 1;
                }

                IsRunning = true;

                DisableWindow();
                UpdateStatus("");

                TokenSources = new List<CancellationTokenSource>();
                Tokens = new List<CancellationToken>();

                for (int i = 0; i < Workers; i++)
                {
                    var tokenSource = new CancellationTokenSource();
                    TokenSources.Add(tokenSource);
                    Tokens.Add(tokenSource.Token);
                }

                int nPreviousCurrent = -1;
                int nCurrent = 0;

                Func<File, int, bool> updateStatus = (file, id) =>
                {
                     if (Tokens[id].IsCancellationRequested)
                     {
                         return false;
                     }

                    nCurrent++;
                    if (nPreviousCurrent < nCurrent)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            UpdateStatus("[" + nCurrent + "/" + _project.Files.Count + "] ");
                        });
                        nPreviousCurrent = nCurrent;
                    }

                    return true;
                 };

                bool bConnect = Workers > 1;
                var partitions = Split(_project.Files, (int)Math.Ceiling(_project.Files.Count / (double)Workers));
                var tasks = new List<Task>();

                for (int i = 0; i < Workers; i++)
                {
                    int id = i;
                    var task = Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            Tokens[id].ThrowIfCancellationRequested();
                            GetTags(updateStatus, partitions[id], id, bConnect);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            Debug.WriteLine(ex.StackTrace);
                        }
                    }, Tokens[id]);
                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);

                Dispatcher.Invoke(() =>
                {
                    EnableWindow();
                    UpdateStatus("");
                });

                for (int i = 0; i < Workers; i++)
                {
                    TokenSources[i].Dispose();
                }

                IsRunning = false;
            }
            else
            {
                for (int i = 0; i < Workers; i++)
                {
                    TokenSources[i].Cancel();
                    TokenSources[i].Dispose();
                }
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

        private class Sheet
        {
            public string Key;
            public TagSet TagSet;
            public Element<string>[] Elements;
            public int nRows;
            public int nColumns;
            public object[,] Values;
        }

        private void ToValues(Sheet sheet)
        {
            int nTagDefinitions = sheet.TagSet.TagDefinitions.Count;

            sheet.nRows = sheet.Elements.Length + 1;
            sheet.nColumns = (2 * nTagDefinitions) + 1;
            sheet.Values = new object[sheet.nRows, sheet.nColumns];

            for (int i = 0; i < sheet.TagSet.TagDefinitions.Count; i++)
            {
                sheet.Values[0, i] = sheet.TagSet.TagDefinitions[i].Name;
                sheet.Values[0, nTagDefinitions + i] = "ID_" + sheet.TagSet.TagDefinitions[i].Name;
            }
            sheet.Values[0, 2 * nTagDefinitions] = "Path";

            for (int i = 0; i < sheet.Elements.Length; i++)
            {
                var element = sheet.Elements[i];
                for (int j = 0; j < nTagDefinitions; j++)
                {
                    var tagDefinition = sheet.TagSet.TagDefinitions[j];
                    foreach (var tag in element.Tags)
                    {
                        if (tag.TagDefinitionName == tagDefinition.Name)
                        {
                            sheet.Values[i + 1, j] = tag.Value;
                            sheet.Values[i + 1, nTagDefinitions + j] = tag.ID.ToString();
                            break;
                        }
                    }
                }
                sheet.Values[i + 1, 2 * nTagDefinitions] = element.File.Path;
            }
        }

        private void ExportTags()
        {
            try
            {
                // create tags values

                object[,] tagValues;
                Tag[] tags = _project.Files.SelectMany(f => f.Tags).ToArray();
                ToValues(tags, out tagValues);

                // create elements values

                TagSet[] tagSets = _project.Files.SelectMany(f => f.TagSets).ToArray();

                Sheet[] sheets = _project.Files.SelectMany(f => f.ElementsByTagSet)
                                               .GroupBy(e => e.Key).Select(g =>
                                                {
                                                    return new Sheet
                                                    {
                                                        Key = g.Key,
                                                        TagSet = tagSets.FirstOrDefault(ts => ts.Name == g.Key),
                                                        Elements = g.ToArray()
                                                    };
                                                }).ToArray();

                foreach (var sheet in sheets)
                {
                    ToValues(sheet);
                }

                // create excel worksheets

                using (var excel = new Excelnterop())
                {
                    excel.CreateApplication();
                    excel.CreateWorkbook();

                    excel.ExportValues(tagValues, tags.Length + 1, 6, "Tags");

                    foreach (var sheet in sheets)
                    {
                        excel.ExportValues(sheet.Values, sheet.nRows, sheet.nColumns, sheet.TagSet.Name);
                    }
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

        private async void TagsGet_Click(object sender, RoutedEventArgs e)
        {
            await GetTags();
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
