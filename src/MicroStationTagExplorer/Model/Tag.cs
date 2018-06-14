using BCOM = MicroStationDGN;

namespace MicroStationTagExplorer.Model
{
    public class Tag
    {
        public string TagSetName { get; set; }
        public string TagDefinitionName { get; set; }
        public object Value { get; set; }
        public BCOM.DLong ID { get; set; }
        public BCOM.DLong HostID { get; set; }
    }
}
