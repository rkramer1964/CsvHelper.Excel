using CsvHelper.Configuration;
using PKCsvHelper.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace PKCsvHelper.Excel
{
    public class FormattedExcelParser : ExcelParserBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public FormattedExcelParser(string path) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), null, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public FormattedExcelParser(string path, string sheetName) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public FormattedExcelParser(string path, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), null, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public FormattedExcelParser(string path, string sheetName, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="FormattedExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public FormattedExcelParser(Stream stream, CultureInfo culture, bool leaveOpen = false) : this(stream, null, culture,
            leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="FormattedExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public FormattedExcelParser(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false) : this(stream,
            sheetName, new CsvConfiguration(culture) { LeaveOpen = leaveOpen })
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="path">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public FormattedExcelParser(string path, string sheetName, CsvConfiguration configuration) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, configuration)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormattedExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public FormattedExcelParser(Stream stream, string sheetName, CsvConfiguration configuration) : base(stream, sheetName, configuration)
        {
        }

        protected override string[] GetRecord()
        {
            var currentRow = Worksheet.Row(Row);
            var cells = currentRow.Cells(1, Count);
            var values = Configuration.TrimOptions.HasFlag(TrimOptions.Trim)
                ? cells.Select(x => x.GetFormattedString()?.Trim()).ToArray()
                : cells.Select(x => x.GetFormattedString()).ToArray();

            return values;
        }
    }
}
