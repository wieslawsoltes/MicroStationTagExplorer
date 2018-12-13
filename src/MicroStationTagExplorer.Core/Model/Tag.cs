using System;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer.Model
{
    [DataContract(IsReference = false, Name = "Tag"), XmlRoot("Tag")]
    public class Tag
    {
        [DataMember(Name = "TagSetName"), XmlAttribute("TagSetName")]
        public string TagSetName { get; set; }

        [DataMember(Name = "TagDefinitionName"), XmlAttribute("TagDefinitionName")]
        public string TagDefinitionName { get; set; }

        [DataMember(Name = "Value"), XmlElement("Value")]
        public object Value { get; set; }

        [DataMember(Name = "ID"), XmlAttribute("ID")]
        public Int64 ID { get; set; }

        [DataMember(Name = "HostID"), XmlAttribute("HostID")]
        public Int64 HostID { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public File File { get; set; }
    }
}
