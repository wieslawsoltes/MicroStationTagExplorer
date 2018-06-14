using System.Collections.Generic;

namespace MicroStationTagExplorer
{
    public class File
    {
        public string Path { get; set; }
        public IList<TagSet> TagSets { get; set; }
        public IList<Tag> Tags { get; set; }
    }
}
