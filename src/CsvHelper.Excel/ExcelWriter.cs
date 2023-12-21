using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;

#pragma warning disable 649
#pragma warning disable 169

namespace PKCsvHelper.Excel
{
    /// <summary>
    /// Used to write CSV files.
    /// </summary>
    public class ExcelWriter : ExcelWriterBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public ExcelWriter(string path) : this(File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), "export",
            CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), "export", culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public ExcelWriter(string path, string sheetName) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, string sheetName, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen = false) : this(stream, "export", culture,
            leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false) : this(stream,
            sheetName, new CsvConfiguration(culture) { LeaveOpen = leaveOpen })
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelWriter(Stream stream, string sheetName, CsvConfiguration configuration) : base(stream, sheetName, configuration) { }

        /// <inheritdoc/>
        protected override void WriteToCell(string value)
        {
            var length = value?.Length ?? 0;

            if (value == null || length == 0)
            {
                return;
            }

            Worksheet.Worksheet.AsRange().Cell(Row, Index).Value = value;
        }
    }
}