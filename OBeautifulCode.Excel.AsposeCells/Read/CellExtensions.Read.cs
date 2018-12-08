// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Extensions methods on type <see cref="Cell"/>.
    /// </summary>
    public static partial class CellExtensions
    {
        /// <summary>
        /// Converts a cell to a range.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The range equivalent to the specified cell.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static Range GetRange(
            this Cell cell)
        {
            new { cell }.Must().NotBeNull();

            var result = cell.Worksheet.GetRange(cell.Row + 1, cell.Row + 1, cell.Column + 1, cell.Column + 1);
            return result;
        }
    }
}
