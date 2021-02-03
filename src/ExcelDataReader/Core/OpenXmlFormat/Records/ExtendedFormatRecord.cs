using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.OpenXmlFormat.Records
{
    internal sealed class ExtendedFormatRecord : Record
    {
        public ExtendedFormatRecord(ExtendedFormat extendedFormat) 
        {
            ExtendedFormat = extendedFormat;
        }

        public ExtendedFormat ExtendedFormat { get; }
    }
}
