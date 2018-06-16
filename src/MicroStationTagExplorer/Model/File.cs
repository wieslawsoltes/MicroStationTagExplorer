using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [XmlRoot("File")]
    public class File
    {
        [XmlAttribute("Path")]
        public string Path { get; set; }

        [XmlArray("TagSets")]
        public ObservableCollection<TagSet> TagSets { get; set; }

        [XmlArray("Tags")]
        public ObservableCollection<Tag> Tags { get; set; }

        [XmlIgnore]
        public IEnumerable<IGrouping<Int64, Tag>> ElementsByHostID { get; set; }

        [XmlIgnore]
        public IEnumerable<IGrouping<string, Tag>> ElementsByTagSet { get; set; }

        [XmlIgnore]
        public IEnumerable<string> Errors { get; set; }
    }
}
