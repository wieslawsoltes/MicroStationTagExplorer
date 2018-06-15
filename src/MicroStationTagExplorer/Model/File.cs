using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace MicroStationTagExplorer
{
    public class File
    {
        public string Path { get; set; }
        public ObservableCollection<TagSet> TagSets { get; set; }
        public ObservableCollection<Tag> Tags { get; set; }
    }
}
