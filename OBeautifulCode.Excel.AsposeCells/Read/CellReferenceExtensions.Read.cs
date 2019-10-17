﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReferenceExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using OBeautifulCode.Assertion.Recipes;

    /// <summary>
    /// Extensions methods on type <see cref="CellReference"/>.
    /// </summary>
    public static partial class CellReferenceExtensions
    {
        /// <summary>
        /// Gets a cell by it's reference.
        /// </summary>
        /// <param name="cellReference">The cell reference.</param>
        /// <param name="workbook">The workbook.</param>
        /// <returns>
        /// The cell corresponding to the specified reference.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cellReference"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static Cell GetCell(
            this CellReference cellReference,
            Workbook workbook)
        {
            new { cellReference }.AsArg().Must().NotBeNull();
            new { workbook }.AsArg().Must().NotBeNull();

            var result = workbook.Worksheets[cellReference.WorksheetName].Cells[cellReference.RowNumber - 1, cellReference.ColumnNumber - 1];

            return result;
        }
    }
}
