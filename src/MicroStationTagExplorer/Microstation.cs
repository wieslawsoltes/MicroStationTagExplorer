using System;
using System.Collections.Generic;
using MicroStationTagExplorer.Model;
using BCOM = MicroStationDGN;

namespace MicroStationTagExplorer
{
    public static class Microstation
    {
        public static Int64 ToInt64(BCOM.DLong value)
        {
            return (Int64)value.High << 32 | (Int64)(UInt32)value.Low;
        }

        public static BCOM.DLong ToInt64(Int64 value)
        {
            return new BCOM.DLong()
            {
                High = (Int32)(value & UInt32.MaxValue),
                Low = (Int32)(value >> 32)
            };
        }

        public static void GetTagData(File file)
        {
            BCOM.Application app = Utilities.CreateObject<BCOM.Application>("MicroStationDGN.Application");
            app.Visible = true;

            BCOM.DesignFile designFile = app.OpenDesignFile(file.Path);

            foreach (BCOM.ModelReference model in designFile.Models)
            {
                if (model.Type == BCOM.MsdModelType.msdModelTypeNormal)
                {
                    model.Activate();
                    break;
                }
            }

            file.TagSets = GetTagSets(designFile);
            file.Tags = GetTags(app.ActiveModelReference, file.Path);
        }

        public static IList<TagSet> GetTagSets(BCOM.DesignFile designFile)
        {
            var tagSets = new List<TagSet>();

            foreach (BCOM.TagSet ts in designFile.TagSets)
            {
                var tagSet = new TagSet()
                {
                    Name = ts.Name,
                    TagDefinitions = new List<TagDefinition>()
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

        public static IList<Tag> GetTags(BCOM.ModelReference model, string path)
        {
            BCOM.ElementScanCriteria sc = Utilities.CreateObject<BCOM.ElementScanCriteria>("MicroStationDGN.ElementScanCriteria");
            sc.ExcludeAllTypes();
            sc.IncludeType(BCOM.MsdElementType.msdElementTypeTag);

            BCOM.ElementEnumerator ee = model.Scan(sc);
            Array elements = ee.BuildArrayFromContents();

            var tags = new List<Tag>();

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
                    Path = path
                };
                tags.Add(tag);
            }

            return tags;
        }
    }
}
