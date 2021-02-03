using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.OpenXmlFormat.Records
{
    internal sealed class ColumnRecord : Record
    {
        public ColumnRecord(Column column)
        {
            Column = column;
        }

        public Column Column { get; }
    }
}
