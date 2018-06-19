using System.Collections.ObjectModel;

namespace MicroStationTagExplorer
{
    public class Sheet
    {
        public bool IsExported { get; set; }
        public string Name { get; set; }
        public TagSet TagSet { get; set; }
        public ObservableCollection<Element<string>> Elements { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        public object[,] Values { get; set; }
    }
}
