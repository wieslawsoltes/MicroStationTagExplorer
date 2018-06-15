using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("Project")]
    public class Project
    {
        [XmlArray("Files")]
        public ObservableCollection<File> Files { get; set; }
    }
}
