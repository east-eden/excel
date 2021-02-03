using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Excel.Core.OpenXmlFormat.Records;

namespace Excel.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlSharedStringsReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string ElementSst = "sst";
        private const string ElementStringItem = "si";

        public XmlSharedStringsReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementSst, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementStringItem, NsSpreadsheetMl))
                {
                    var value = StringHelper.ReadStringItem(Reader);
                    yield return new SharedStringRecord(value);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }
    }
}
