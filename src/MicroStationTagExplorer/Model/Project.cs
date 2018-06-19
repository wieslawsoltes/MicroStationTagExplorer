﻿using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace MicroStationTagExplorer
{
    [DataContract(IsReference = false, Name = "Project"), XmlRoot("Project")]
    public class Project
    {
        [DataMember(Name = "Name"), XmlAttribute("Name")]
        public string Name { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public string Path { get; set; }

        [DataMember(Name = "Files"), XmlArray("Files")]
        public ObservableCollection<File> Files { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Tag> Tags { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public object[,] TagValues { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<TagSet> TagSets { get; set; }

        [IgnoreDataMember, XmlIgnore]
        public ObservableCollection<Sheet> Sheets { get; set; }
    }
}
