using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = true, Name = "File"), XmlRoot("File")]
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

        [IgnoreDataMember, XmlIgnore]
        public IEnumerable<IGrouping<Int64, Tag>> ElementsByHostID { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public IEnumerable<IGrouping<string, Tag>> ElementsByTagSet { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public IEnumerable<Error> Errors { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public bool HasErrors { get; set; }
    }
}
