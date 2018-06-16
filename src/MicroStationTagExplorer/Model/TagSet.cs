using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = true, Name = "TagSet"), XmlRoot("TagSet")]
    public class TagSet
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }

        [DataMember(Name = "TagDefinitions"), XmlArray("TagDefinitions")]
        public ObservableCollection<TagDefinition> TagDefinitions { get; set; }

        [DataMember(Name = "File"), XmlAttribute("Name")]
        public File File { get; set; }
    }
}
