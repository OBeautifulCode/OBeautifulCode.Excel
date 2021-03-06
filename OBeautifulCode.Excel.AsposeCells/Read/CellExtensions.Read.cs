﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using MoreLinq;

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
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = cell.Worksheet.GetRange(cell.GetRowNumber(), cell.GetRowNumber(), cell.GetColumnNumber(), cell.GetColumnNumber());
            return result;
        }

        /// <summary>
        /// Gets the width of a cell, in pixels, accounting for merged cells.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="includeMergedCells">Optional value indicating whether to include other cells that are merged with the specified cell.  Default is to include merged cells.</param>
        /// <returns>
        /// The width of the cell in pixels.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static int GetWidthInPixels(
            this Cell cell,
            bool includeMergedCells = true)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = 0;

            if (includeMergedCells && cell.IsMerged)
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

        /// <summary>
        /// Gets the height of a cell, in pixels, accounting for merged cells.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="includeMergedCells">Optional value indicating whether to include other cells that are merged with the specified cell.  Default is to include merged cells.</param>
        /// <returns>
        /// The height of the cell in pixels.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static int GetHeightInPixels(
            this Cell cell,
            bool includeMergedCells = true)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = 0;

            if (includeMergedCells && cell.IsMerged)
            {
                var mergedRange = cell.GetMergedRange();
                var cellsWithDistinctRows = mergedRange.GetCells().DistinctBy(_ => _.Row);

                foreach (var cellWithDistinctRow in cellsWithDistinctRows)
                {
                    result = result + cell.Worksheet.Cells.GetRowHeightPixel(cellWithDistinctRow.Row);
                }
            }
            else
            {
                result = result + cell.Worksheet.Cells.GetRowHeightPixel(cell.Row);
            }

            return result;
        }

        /// <summary>
        /// Gets the cell's row number.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The cell's row number.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static int GetRowNumber(
            this Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = cell.Row + 1;
            return result;
        }

        /// <summary>
        /// Gets the cell's column number.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The cell's column number.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static int GetColumnNumber(
            this Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = cell.Column + 1;
            return result;
        }

        /// <summary>
        /// Gets the <see cref="CellReference"/> corresponding to the specified cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The cell reference corresponding to the specified cell.
        /// </returns>
        public static CellReference ToCellReference(
            this Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            var result = new CellReference(cell.Worksheet.Name, cell.Row + 1, cell.Column + 1);

            return result;
        }
    }
}
