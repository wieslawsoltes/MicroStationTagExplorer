using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using BCOM = MicroStationDGN;

namespace MicroStationTagExplorer
{
    internal class MicrostationInterop : IDisposable
    {
        private string _path;
        private BCOM.Application _application;
        private BCOM.DesignFile _designFile;

        private static Int64 ToInt64(BCOM.DLong value)
        {
            return (Int64)value.High << 32 | (Int64)(UInt32)value.Low;
        }

        private static BCOM.DLong ToInt64(Int64 value)
        {
            return new BCOM.DLong()
            {
                High = (Int32)(value & UInt32.MaxValue),
                Low = (Int32)(value >> 32)
            };
        }

        public MicrostationInterop(string path)
        {
            Initialize(path);
        }

        private void Initialize(string path)
        {
            _path = path;
            _application = ComUtilities.CreateObject<BCOM.Application>("MicroStationDGN.Application");
            _application.Visible = true;
            _designFile = _application.OpenDesignFile(_path);
        }

        public void SetNormalActiveModel()
        {
            foreach (BCOM.ModelReference model in _designFile.Models)
            {
                if (model.Type == BCOM.MsdModelType.msdModelTypeNormal)
                {
                    model.Activate();
                    break;
                }
            }
        }

        public ObservableCollection<TagSet> GetTagSets()
        {
            var tagSets = new ObservableCollection<TagSet>();

            foreach (BCOM.TagSet ts in _designFile.TagSets)
            {
                var tagSet = new TagSet()
                {
                    Name = ts.Name,
                    TagDefinitions = new ObservableCollection<TagDefinition>()
                };
                tagSets.Add(tagSet);

                foreach (BCOM.TagDefinition td in ts.TagDefinitions)
                {
                    var tagDefinition = new TagDefinition()
                    {
                        Name = td.Name
                    };
                    tagSet.TagDefinitions.Add(tagDefinition);
                }
            }

            return tagSets;
        }

        public ObservableCollection<Tag> GetTags()
        {
            var tags = new ObservableCollection<Tag>();

            BCOM.ElementScanCriteria sc = ComUtilities.CreateObject<BCOM.ElementScanCriteria>("MicroStationDGN.ElementScanCriteria");
            sc.ExcludeAllTypes();
            sc.IncludeType(BCOM.MsdElementType.msdElementTypeTag);

            BCOM.ElementEnumerator ee = _application.ActiveModelReference.Scan(sc);
            Array elements = ee.BuildArrayFromContents();

            foreach (BCOM.Element element in elements)
            {
                BCOM.TagElement te = element as BCOM.TagElement;

                var tag = new Tag()
                {
                    TagSetName = te.TagSetName,
                    TagDefinitionName = te.TagDefinitionName,
                    Value = te.Value,
                    ID = ToInt64(te.ID),
                    HostID = ToInt64(te.BaseElement != null ? te.BaseElement.ID : new BCOM.DLong()),
                    Path = _path
                };
                tags.Add(tag);
            }

            return tags;
        }

        public void Dispose()
        {
            ComUtilities.ReleaseComObject(_designFile);
            _designFile = null;

            //if (_application != null)
            //{
            //    _application.Quit();
            //}

            ComUtilities.ReleaseComObject(_application);
            _application = null;
        }
    }
}
