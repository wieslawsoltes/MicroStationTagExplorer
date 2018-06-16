using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = true, Name = "Project"), XmlRoot("Project")]
    public class Project
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }

        [DataMember(Name = "Files"), XmlArray("Files")]
        public ObservableCollection<File> Files { get; set; }
    }
}
