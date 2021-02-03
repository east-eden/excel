using System.IO;
using System.Text;
using Excel.Core.CsvFormat;

namespace Excel
{
    internal class ExcelCsvReader : Excel<CsvWorkbook, CsvWorksheet>
    {
        public ExcelCsvReader(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators, int analyzeInitialCsvRows)
        {
            Workbook = new CsvWorkbook(stream, fallbackEncoding, autodetectSeparators, analyzeInitialCsvRows);

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
