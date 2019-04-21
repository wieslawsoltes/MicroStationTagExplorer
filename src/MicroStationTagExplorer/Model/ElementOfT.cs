using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer.Core.Model
{
    [DataContract(IsReference = false, Name = "Element"), XmlRoot("Element")]
    public class Element<T> : Element
    {
        [DataMember(Name = "Key"), XmlAttribute("Key")]
        public T Key { get; set; }
    }
}
