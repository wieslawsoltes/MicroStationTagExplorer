using System;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = false, Name = "File"), XmlRoot("File")]
    public class File
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }

        [DataMember(Name = "Path"), XmlAttribute("Path")]
        public string Path { get; set; }

        [DataMember(Name = "TagSets"), XmlArray("TagSets")]
        public ObservableCollection<TagSet> TagSets { get; set; }

        [DataMember(Name = "Tags"), XmlArray("Tags")]
        public ObservableCollection<Tag> Tags { get; set; }

        [DataMember(Name = "Texts"), XmlArray("Texts")]
        public ObservableCollection<Text> Texts { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Element<Int64>> ElementsByHostID { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Element<string>> ElementsByTagSet { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Error> Errors { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public bool HasErrors { get; set; }
    }
}
