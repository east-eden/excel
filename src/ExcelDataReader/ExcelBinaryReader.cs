using System.IO;
using System.Text;
using Excel.Core.BinaryFormat;

namespace Excel
{
    /// <summary>
    /// Excel Class
    /// </summary>
    internal class ExcelBinaryReader : Excel<XlsWorkbook, XlsWorksheet>
    {
        public ExcelBinaryReader(Stream stream, string password, Encoding fallbackEncoding)
        {
            Workbook = new XlsWorkbook(stream, password, fallbackEncoding);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        public override void Close()
        {
            base.Close();
            Workbook?.Stream?.Dispose();
            Workbook = null;
        }
    }
}
