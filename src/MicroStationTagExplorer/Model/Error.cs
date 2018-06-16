using System.Linq;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = true, Name = "Error"), XmlRoot("Error")]
    public class Error
    {
        [DataMember(Name = "Message"), XmlAttribute("Message")]
        public string Message { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public IGrouping<string, Tag> Element { get; set; }

        [DataMember(Name = "TagSet"), XmlAttribute("TagSet")]
        public TagSet TagSet { get; set; }
    }
}
