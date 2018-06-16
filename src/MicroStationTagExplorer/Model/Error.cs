using System.Linq;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("Error")]
    public class Error
    {
        [XmlIgnore]
        public string Message { get; set; }

        [XmlIgnore]
        public IGrouping<string, Tag> Element { get; set; }

        [XmlIgnore]
        public TagSet TagSet { get; set; }
    }
}
