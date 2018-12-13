using System;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer.Model
{
    [DataContract(IsReference = false, Name = "Text"), XmlRoot("Text")]
    public class Text
    {
        [DataMember(Name = "Value"), XmlAttribute("Value")]
        public string Value { get; set; }

        [DataMember(Name = "ID"), XmlAttribute("ID")]
        public Int64 ID { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public File File { get; set; }
    }
}
