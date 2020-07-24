// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using Aspose.Cells;

    using static System.FormattableString;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Extensions methods on type <see cref="Worksheet"/>.
    /// </summary>
    public static partial class WorksheetExtensions
    {
        /// <summary>
        /// Gets the specified range.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="startRowNumber">The row number of the start of the range.</param>
        /// <param name="endRowNumber">The row number of the end of the range.</param>
        /// <param name="startColumnNumber">The column number of the start of the range.</param>
        /// <param name="endColumnNumber">The column number of the end of the range.</param>
        /// <returns>
        /// The range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="startRowNumber"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="startColumnNumber"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="endRowNumber"/> is less than <paramref name="startRowNumber"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="endColumnNumber"/> is less than  <paramref name="startColumnNumber"/>.</exception>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "startRowNumber-1", Justification = "Overflow is not possible based on parameter validation.")]
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "startColumnNumber-1", Justification = "Overflow is not possible based on parameter validation.")]
        public static Range GetRange(
            this Worksheet worksheet,
            int startRowNumber,
            int endRowNumber,
            int startColumnNumber,
            int endColumnNumber)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (startRowNumber < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(startRowNumber)}' < '{1}'"), (Exception)null);
            }

            if (startColumnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(startColumnNumber)}' < '{1}'"), (Exception)null);
            }

            if (endRowNumber < startRowNumber)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(endRowNumber)}' < '{startRowNumber}'"), (Exception)null);
            }

            if (endColumnNumber < startColumnNumber)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(endColumnNumber)}' < '{startColumnNumber}'"), (Exception)null);
            }

            var result = worksheet.Cells.CreateRange(startRowNumber - 1, startColumnNumber - 1, endRowNumber - startRowNumber + 1, endColumnNumber - startColumnNumber + 1);
            return result;
        }

        /// <summary>
        /// Gets a cell.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="columnNumber">The column number.</param>
        /// <returns>
        /// The cell.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="rowNumber"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnNumber"/> is less than 1.</exception>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "rowNumber-1", Justification = "Overflow is not possible based on parameter validation.")]
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "columnNumber-1", Justification = "Overflow is not possible based on parameter validation.")]
        public static Cell GetCell(
            this Worksheet worksheet,
            int rowNumber,
            int columnNumber)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (rowNumber < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(rowNumber)}' < '{1}'"), (Exception)null);
            }

            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnNumber)}' < '{1}'"), (Exception)null);
            }

            var result = worksheet.Cells[rowNumber - 1, columnNumber - 1];
            return result;
        }
    }
}
