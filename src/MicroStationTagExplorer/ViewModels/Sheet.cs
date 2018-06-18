
namespace MicroStationTagExplorer
{
    public class Sheet
    {
        public string Key;
        public TagSet TagSet;
        public Element<string>[] Elements;
        public int nRows;
        public int nColumns;
        public object[,] Values;
    }
}
