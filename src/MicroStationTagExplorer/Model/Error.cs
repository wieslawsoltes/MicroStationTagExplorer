using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = false, Name = "Error"), XmlRoot("Error")]
    public class Error
    {
        [DataMember(Name = "Message"), XmlAttribute("Message")]
        public string Message { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public Element Element { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public TagSet TagSet { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public File File { get; set; }
    }
}
