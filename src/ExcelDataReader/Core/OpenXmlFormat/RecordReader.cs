using System;
using System.Collections.Generic;
using System.Text;

using Excel.Core.OpenXmlFormat.Records;

#nullable enable

namespace Excel.Core.OpenXmlFormat
{
    internal abstract class RecordReader : IDisposable
    {
        ~RecordReader()
        {
            Dispose(false);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public abstract Record? Read();

        protected virtual void Dispose(bool disposing) 
        {
        }
    }
}
