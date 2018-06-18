using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Serialization;
using BCOM = MicroStationDGN;

namespace MicroStationTagExplorer
{
    public class TagExplorer
    {
        private static StringComparison _comparisonType = StringComparison.OrdinalIgnoreCase;
        private static string _xmlExt = ".xml";
        private static string _dgnExt = ".dgn";
        private static string _dwgExt = ".dwg";

        public volatile bool IsRunning;
        public List<CancellationTokenSource> TokenSources { get; set; }
        public List<CancellationToken> Tokens { get; set; }
        public int WorkersNum { get; set; }
        public List<Worker> Workers { get; set; }
        public List<Worker> ActiveWorkers { get; set; }
        public Project Project { get; set; }

        public TagExplorer()
        {
            IsRunning = false;
            WorkersNum = 1;
            Workers = new List<Worker>();
            ActiveWorkers = new List<Worker>();
        }

        public void ValidateFile(File file)
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

        public void UpdateProject(Project project)
        {
            foreach (var file in project.Files)
            {
                ValidateFile(file);
            }
        }

        public IEnumerable<Error> ValidateTags(File file)
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

        public IEnumerable<Error> ValidateTags(Element element, TagSet tagSet)
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

        public bool IsProject(string path)
        {
            return path.EndsWith(_xmlExt, _comparisonType);
        }

        public bool IsSupportedFile(string path)
        {
            return path.EndsWith(_dwgExt, _comparisonType) || path.EndsWith(_dgnExt, _comparisonType);
        }

        public void AddFile(string path)
        {
            if (IsSupportedFile(path))
            {
                var file = new File()
                {
                    Name = System.IO.Path.GetFileName(path),
                    Path = path
                };
                Project.Files.Add(file);
            }
        }

        public void AddFiles(string[] fileNames)
        {
            foreach (var fileName in fileNames)
            {
                AddFile(fileName);
            }
        }

        public void AddPaths(string[] paths)
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

        public void NewProject()
        {
            Project = new Project()
            {
                Name = "project",
                Path = "project" + _xmlExt,
                Files = new ObservableCollection<File>()
            };

            UpdateProject(Project);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void SerializeXmlSerializer(Project project, string fileName)
        {
            using (var writer = new System.IO.StreamWriter(fileName))
            {
                var serializer = new XmlSerializer(typeof(Project));
                serializer.Serialize(writer, project);
            }
        }

        public Project DeserializeXmlSerializer(string fileName)
        {
            Project project = null;
            using (var reader = new System.IO.StreamReader(fileName))
            {
                var serializer = new XmlSerializer(typeof(Project));
                project = (Project)serializer.Deserialize(reader);
            }
            return project;
        }

        public void SerializeDataContract(Project project, string fileName)
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

        public Project DeserializeDataContract(string fileName)
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

        public void OpenProject(string fileName)
        {
            Project project = DeserializeXmlSerializer(fileName);
            if (project != null)
            {
                Project = project;
                Project.Path = fileName;

                foreach (var file in Project.Files)
                {
                    file.Name = System.IO.Path.GetFileName(file.Path);
                }

                UpdateProject(Project);
            }
        }

        public void SaveProject(string fileName)
        {
            SerializeXmlSerializer(Project, fileName);
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

        public void GetTags(Func<File, int, bool> updateStatus, IEnumerable<File> files, int id, bool bConnect, Worker worker)
        {
            using (var microstation = new MicrostationInterop())
            {
                if (worker == null)
                {
                    if (bConnect == true)
                    {
                        microstation.ConnectApplication();
                    }
                    else
                    {
                        microstation.CreateApplication();
                    }
                }
                else
                {
                    microstation.SetApplication(worker.Application);
                }

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

        public void ToValues(Tag[] tags, out object[,] values)
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

        public void ToValues(Sheet sheet)
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

        public void ImportTags()
        {
        }

        public void ExportTags()
        {
            // create tags values

            object[,] tagValues;
            Tag[] tags = Project.Files.SelectMany(f => f.Tags).ToArray();
            ToValues(tags, out tagValues);

            // create elements values

            TagSet[] tagSets = Project.Files.SelectMany(f => f.TagSets).ToArray();

            Sheet[] sheets = Project.Files.SelectMany(f => f.ElementsByTagSet)
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

        public void GetWorkers()
        {
            string[] progIDs = { "MicroStationDGN.Application" };
            var results = ComUtilities.GetRunningCOMObjects(progIDs);
            foreach (var result in results)
            {
                BCOM.Application application = result.RunningObject as BCOM.Application;
                if (application != null)
                {
                    var worker = new Worker()
                    {
                        IsEnabled = true,
                        Result = result,
                        Application = application
                    };
                    Workers.Add(worker);
                }
            }
        }
    }
}
