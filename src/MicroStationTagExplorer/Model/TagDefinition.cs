
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("TagDefinition")]
    public class TagDefinition
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }
    }
}
