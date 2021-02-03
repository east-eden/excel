using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.OpenXmlFormat.Records
{
    internal sealed class MergeCellRecord : Record
    {
        public MergeCellRecord(CellRange range) 
        {
            Range = range;
        }

        public CellRange Range { get; }
    }
}
