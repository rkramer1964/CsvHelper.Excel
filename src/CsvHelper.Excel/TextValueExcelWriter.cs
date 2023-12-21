using ClosedXML;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace PKCsvHelper.Excel
{
    public class TextValueExcelWriter : ExcelWriter
    {
        public TextValueExcelWriter(string path) : base(path)
        {
        }

        public TextValueExcelWriter(string path, CultureInfo culture) : base(path, culture)
        {
        }

        public TextValueExcelWriter(string path, string sheetName) : base(path, sheetName)
        {
        }

        public TextValueExcelWriter(string path, string sheetName, CultureInfo culture) : base(path, sheetName, culture)
        {
        }

        public TextValueExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen = false) : base(stream, culture, leaveOpen)
        {
        }

        public TextValueExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false) : base(stream, sheetName, culture, leaveOpen)
        {
        }

        protected override void WriteToCell(string value)
        {
            var length = value?.Length ?? 0;

            if (value == null || length == 0)
            {
                return;
            }
            Worksheet.Worksheet.AsRange().Cell(Row, Index).DataType = XLDataType.Text;

            Worksheet.Worksheet.AsRange().Cell(Row, Index).Value = value;
        }
    }
}
