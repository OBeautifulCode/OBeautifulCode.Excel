// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using MoreLinq;

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

        /// <summary>
        /// Gets the width of a cell, in pixels, accounting for merged cells.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The width of the cell in pixels.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static int GetWidthInPixels(
            this Cell cell)
        {
            new { cell }.Must().NotBeNull();

            var result = 0;

            if (cell.IsMerged)
            {
                var mergedRange = cell.GetMergedRange();
                var cellsWithDistinctColumns = mergedRange.GetCells().DistinctBy(_ => _.Column);

                foreach (var cellWithDistinctColumn in cellsWithDistinctColumns)
                {
                    result = result + cell.Worksheet.Cells.GetColumnWidthPixel(cellWithDistinctColumn.Column);
                }
            }
            else
            {
                result = result + cell.Worksheet.Cells.GetColumnWidthPixel(cell.Column);
            }

            return result;
        }
    }
}
