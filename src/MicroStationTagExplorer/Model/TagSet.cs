using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace MicroStationTagExplorer
{
    public class TagSet
    {
        public string Name { get; set; }
        public ObservableCollection<TagDefinition> TagDefinitions { get; set; }
    }
}
