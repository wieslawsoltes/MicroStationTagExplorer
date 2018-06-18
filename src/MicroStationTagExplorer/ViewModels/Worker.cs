using BCOM = MicroStationDGN;

namespace MicroStationTagExplorer
{
    public class Worker
    {
        public bool IsEnabled { get; set; }
        public RunningObjectResult Result { get; set; }
        public BCOM.Application Application { get; set; }
    }
}
