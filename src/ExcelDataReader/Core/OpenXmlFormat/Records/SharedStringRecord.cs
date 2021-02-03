using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.OpenXmlFormat.Records
{
    internal sealed class SharedStringRecord : Record
    {
        public SharedStringRecord(string value) 
        {
            Value = value;
        }

        public string Value { get; }
    }
}
