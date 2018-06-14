using System.Collections.Generic;

namespace MicroStationTagExplorer.Model
{
    public class DgnFile
    {
        public string Path { get; set; }
        public IList<TagSet> TagSets { get; set; }
        public IList<Tag> Tags { get; set; }
    }
}
