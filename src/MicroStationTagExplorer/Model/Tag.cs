using System;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("Tag")]
    public class Tag
    {
        [XmlAttribute("TagSetName")]
        public string TagSetName { get; set; }

        [XmlAttribute("TagDefinitionName")]
        public string TagDefinitionName { get; set; }

        [XmlElement]
        public object Value { get; set; }

        [XmlAttribute("ID")]
        public Int64 ID { get; set; }

        [XmlAttribute("HostID")]
        public Int64 HostID { get; set; }

        [XmlAttribute("Path")]
        public string Path { get; set; }
    }
}
