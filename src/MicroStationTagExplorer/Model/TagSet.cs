using System.Collections.Generic;

namespace MicroStationTagExplorer.Model
{
    public class TagSet
    {
        public string Name { get; set; }
        public IList<TagDefinition> TagDefinitions { get; set; }
    }
}
