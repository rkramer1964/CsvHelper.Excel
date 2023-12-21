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
    public abstract class ExcelWriterBase : CsvWriter
    {
        private readonly bool _leaveOpen;
        private readonly bool _sanitizeForInjection;

        private bool _disposed;
        private int _row = 1;
        private int _index = 1;
        private readonly IXLWorksheet _worksheet;
        private readonly Stream _stream;

        public override int Index => _index;
        public override int Row => _row;
        public virtual IXLWorksheet Worksheet => _worksheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelWriterBase(Stream stream, string sheetName, CsvConfiguration configuration) : base(TextWriter.Null, configuration)
        {
            configuration.Validate();
            _worksheet = new XLWorkbook(XLEventTracking.Disabled).AddWorksheet(sheetName);
            this._stream = stream;

            _leaveOpen = configuration.LeaveOpen;
            _sanitizeForInjection = configuration.SanitizeForInjection;
        }


        /// <inheritdoc/>
        public override void WriteField(string field, bool shouldQuote)
        {
            if (_sanitizeForInjection)
            {
                field = SanitizeForInjection(field);
            }

            WriteToCell(field);
            _index++;
        }

        /// <inheritdoc/>
        public override void NextRecord()
        {
            Flush();
            _index = 1;
            _row++;
        }

        /// <inheritdoc/>
        public override async Task NextRecordAsync()
        {
            await FlushAsync();
            _index = 1;
            _row++;
        }

        /// <inheritdoc/>
        public override void Flush()
        {
            _stream.Flush();
        }

        /// <inheritdoc/>
        public override Task FlushAsync()
        {
            return _stream.FlushAsync();
        }

        protected abstract void WriteToCell(string value);

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            Flush();
            _worksheet.Workbook.SaveAs(_stream);
            _stream.Flush();

            if (disposing)
            {
                // Dispose managed state (managed objects)
                _worksheet.Workbook.Dispose();
                if (!_leaveOpen)
                {
                    _stream.Dispose();
                }
            }

            // Free unmanaged resources (unmanaged objects) and override finalizer
            // Set large fields to null

            _disposed = true;
        }

#if !NET45 && !NET47 && !NETSTANDARD2_0
		/// <inheritdoc/>
		protected override async ValueTask DisposeAsync(bool disposing)
		{
			if (_disposed)
			{
				return;
			}

			await FlushAsync().ConfigureAwait(false);
			_worksheet.Workbook.SaveAs(_stream);
			await _stream.FlushAsync().ConfigureAwait(false);

			if (disposing)
			{
				// Dispose managed state (managed objects)
				_worksheet.Workbook.Dispose();
				if (!_leaveOpen)
				{
					await _stream.DisposeAsync().ConfigureAwait(false);
				}
			}

			// Free unmanaged resources (unmanaged objects) and override finalizer
			// Set large fields to null


			_disposed = true;
		}
#endif
    }
}