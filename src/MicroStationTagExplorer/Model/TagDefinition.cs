using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = true, Name = "TagDefinition"), XmlRoot("TagDefinition")]
    public class TagDefinition
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }
    }
}
