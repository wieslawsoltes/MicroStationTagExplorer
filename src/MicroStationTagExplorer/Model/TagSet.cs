using System.Collections.Generic;

namespace MicroStationTagExplorer
{
    public class TagSet
    {
        public string Name { get; set; }
        public IList<TagDefinition> TagDefinitions { get; set; }
    }
}
