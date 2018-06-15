using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("TagSet")]
    public class TagSet
    {
        [XmlAttribute("Path")]
        public string Name { get; set; }

        [XmlArray("TagDefinitions")]
        public ObservableCollection<TagDefinition> TagDefinitions { get; set; }
    }
}
