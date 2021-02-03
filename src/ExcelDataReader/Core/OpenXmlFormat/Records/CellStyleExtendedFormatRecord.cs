using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.OpenXmlFormat.Records
{
    internal sealed class CellStyleExtendedFormatRecord : Record
    {
        public CellStyleExtendedFormatRecord(ExtendedFormat extendedFormat)
        {
            ExtendedFormat = extendedFormat;
        }

        public ExtendedFormat ExtendedFormat { get; }
    }
}
