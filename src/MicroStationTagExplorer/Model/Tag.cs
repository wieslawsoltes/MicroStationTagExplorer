using System;

namespace MicroStationTagExplorer.Model
{
    public class Tag
    {
        public string TagSetName { get; set; }
        public string TagDefinitionName { get; set; }
        public object Value { get; set; }
        public Int64 ID { get; set; }
        public Int64 HostID { get; set; }
        public string Path { get; set; }
    }
}
