using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace PKCsvHelper.Excel
{
    public abstract class ExcelParserBase : IParser
    {
        private readonly int _lastRow;
        private readonly bool _leaveOpen;
        private readonly Stream _stream;
        private readonly IXLWorksheet _worksheet;
        private string[] _currentRecord;

        private bool _disposed;
        private int _rawRow = 1;
        private int _row = 1;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        protected ExcelParserBase(Stream stream, string sheetName, CsvConfiguration configuration)
        {
            var workbook = new XLWorkbook(stream, XLEventTracking.Disabled);

            _worksheet = string.IsNullOrEmpty(sheetName) ? workbook.Worksheet(1) : workbook.Worksheet(sheetName);

            Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);
            _stream = stream;
            var lastRowUsed = _worksheet.LastRowUsed();
            if (lastRowUsed != null)
            {
                _lastRow = lastRowUsed.RowNumber();

                var cellsUsed = _worksheet.CellsUsed();
                Count = cellsUsed.Max(c => c.Address.ColumnNumber) -
                    cellsUsed.Min(c => c.Address.ColumnNumber) + 1;
            }

            Context = new CsvContext(this);
            if (configuration != null)
            {
                _leaveOpen = configuration.LeaveOpen; // use the csvconfiguration instead of the IParserConfiguration interface as LeaveOpen was removed in CsvHelper 30.0.0
            }
        }


        public string this[int index] => Record.ElementAtOrDefault(index);

        public long ByteCount => -1;
        public long CharCount => -1;
        public IParserConfiguration Configuration { get; }
        public CsvContext Context { get; }
        public int Count { get; }

        public string Delimiter => Configuration.Delimiter;

        public string RawRecord => string.Join(Delimiter, Record);
        public int RawRow => _rawRow;

        public string[] Record => _currentRecord;
        protected IXLWorksheet Worksheet => _worksheet;

        public int Row => _row;


        /// <inheritdoc/>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public bool Read()
        {
            if (Row > _lastRow)
            {
                return false;
            }

            _currentRecord = GetRecord();
            _row++;
            _rawRow++;
            return true;
        }

        public Task<bool> ReadAsync()
        {
            if (Row > _lastRow)
            {
                return Task.FromResult(false);
            }

            _currentRecord = GetRecord();
            _row++;
            _rawRow++;
            return Task.FromResult(true);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Dispose managed state (managed objects)

                if (!_leaveOpen)
                {
                    _stream?.Dispose();
                }
            }

            // Free unmanaged resources (unmanaged objects) and override finalizer
            // Set large fields to null

            _disposed = true;
        }

        protected abstract string[] GetRecord();
    }
}