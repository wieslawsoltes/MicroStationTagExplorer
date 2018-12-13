using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer.Model
{
    [DataContract(IsReference = false, Name = "TagSet"), XmlRoot("TagSet")]
    public class TagSet
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }

        [DataMember(Name = "TagDefinitions"), XmlArray("TagDefinitions")]
        public ObservableCollection<TagDefinition> TagDefinitions { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public File File { get; set; }
    }
}
