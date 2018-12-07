﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetManipulationExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Extensions methods to manipulate <see cref="Worksheet"/>.
    /// </summary>
    public static class WorksheetManipulationExtensions
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
            new { worksheet }.Must().NotBeNull();
            new { startRowNumber }.Must().BeGreaterThanOrEqualTo(1);
            new { startColumnNumber }.Must().BeGreaterThanOrEqualTo(1);
            new { endRowNumber }.Must().BeGreaterThanOrEqualTo(startRowNumber);
            new { endColumnNumber }.Must().BeGreaterThanOrEqualTo(startColumnNumber);

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
            new { worksheet }.Must().NotBeNull();
            new { rowNumber }.Must().BeGreaterThanOrEqualTo(1);
            new { columnNumber }.Must().BeGreaterThanOrEqualTo(1);

            var result = worksheet.Cells[rowNumber - 1, columnNumber - 1];
            return result;
        }
    }
}
