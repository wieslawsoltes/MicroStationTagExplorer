using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    public class Tag
    {
        public string TagSetName { get; set; }
        public string TagDefinitionName { get; set; }
        public object Value { get; set; }
        public MicroStationDGN.DLong ID { get; set; }
        public MicroStationDGN.DLong HostID { get; set; }
    }

    public class TagDefinition
    {
        public string Name { get; set; }
    }

    public class TagSet
    {
        public string Name { get; set; }
        public IList<TagDefinition> TagDefinitions { get; set; }
    }

    public class DgnFile
    {
        public string Path { get; set; }
        public IList<TagSet> TagSets { get; set; }
        public IList<Tag> Tags { get; set; }
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private T CreateObject<T>(string progID)
        {
            return (T)Activator.CreateInstance(Type.GetTypeFromProgID(progID));
        }

        private void GetTagData(DgnFile dgnFile)
        {
            MicroStationDGN.Application app = CreateObject<MicroStationDGN.Application>("MicroStationDGN.Application");
            app.Visible = true;

            MicroStationDGN.DesignFile designFile = app.OpenDesignFile(dgnFile.Path);

            foreach (MicroStationDGN.ModelReference model in designFile.Models)
            {
                if (model.Type == MicroStationDGN.MsdModelType.msdModelTypeNormal)
                {
                    model.Activate();
                    break;
                }
            }

            dgnFile.TagSets = GetTagSets(designFile);
            dgnFile.Tags = GetTags(app.ActiveModelReference);
        }

        private IList<TagSet> GetTagSets(MicroStationDGN.DesignFile designFile)
        {
            var tagSets = new List<TagSet>();

            foreach (MicroStationDGN.TagSet ts in designFile.TagSets)
            {
                var tagSet = new TagSet()
                {
                    Name = ts.Name,
                    TagDefinitions = new List<TagDefinition>()
                };
                tagSets.Add(tagSet);

                foreach (MicroStationDGN.TagDefinition td in ts.TagDefinitions)
                {
                    var tagDefinition = new TagDefinition()
                    {
                        Name = td.Name
                    };
                    tagSet.TagDefinitions.Add(tagDefinition);
                }
            }

            return tagSets;
        }

        private IList<Tag> GetTags(MicroStationDGN.ModelReference model)
        {
            MicroStationDGN.ElementScanCriteria sc = CreateObject<MicroStationDGN.ElementScanCriteria>("MicroStationDGN.ElementScanCriteria");
            sc.ExcludeAllTypes();
            sc.IncludeType(MicroStationDGN.MsdElementType.msdElementTypeTag);

            MicroStationDGN.ElementEnumerator ee = model.Scan(sc);
            Array elements = ee.BuildArrayFromContents();

            var tags = new List<Tag>();

            foreach (MicroStationDGN.Element element in elements)
            {
                MicroStationDGN.TagElement te = element as MicroStationDGN.TagElement;

                var tag = new Tag()
                {
                    TagSetName = te.TagSetName,
                    TagDefinitionName = te.TagDefinitionName,
                    Value = te.Value,
                    ID = te.ID,
                    HostID = te.BaseElement != null ? te.BaseElement.ID : new MicroStationDGN.DLong(),
                };
                tags.Add(tag);
            }

            return tags;
        }

        private void FileAdd_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "ALl Files (*.*)|*.*",
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

        private void ButtonGetData_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridFiles.DataContext != null)
            {
                int selectedIndex = DataGridFiles.SelectedIndex;
                IList<DgnFile> dgnFiles = DataGridFiles.DataContext as IList<DgnFile>;
                if (dgnFiles != null)
                {
                    ButtonGetData.IsEnabled = false;

                    Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            Dispatcher.Invoke(() => TextBlockStatus.Text = "");
                            foreach (var dgnFile in dgnFiles)
                            {
                                Dispatcher.Invoke(() => TextBlockStatus.Text = dgnFile.Path);
                                GetTagData(dgnFile);
                            }
                            Dispatcher.Invoke(() => TextBlockStatus.Text = "");
                            Dispatcher.Invoke(() => ButtonGetData.IsEnabled = true);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            Debug.WriteLine(ex.StackTrace);
                        }
                    });
                }
            }
        }

        private void ButtonExportData_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application app = CreateObject<Excel.Application>("Excel.Application");
            app.Visible = true;

            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets.Add();
        }
    }
}
