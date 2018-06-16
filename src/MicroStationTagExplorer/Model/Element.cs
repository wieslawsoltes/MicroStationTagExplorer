using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = false, Name = "Element"), XmlRoot("Element")]
    public abstract class Element
    {
        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Tag> Tags { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Error> Errors { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public bool HasErrors { get; set; }
    }
}
