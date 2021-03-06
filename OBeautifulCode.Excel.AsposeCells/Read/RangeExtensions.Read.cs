﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeExtensions.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using Aspose.Cells;

    using static System.FormattableString;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Extensions methods on type <see cref="Range"/>.
    /// </summary>
    public static partial class RangeExtensions
    {
        /// <summary>
        /// Gets the row numbers in the specified range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The row numbers in the range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static IReadOnlyList<int> GetRowNumbers(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var result = Enumerable.Range(range.FirstRow + 1, range.RowCount).ToList();

            return result;
        }

        /// <summary>
        /// Gets the column numbers in the specified range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The column numbers in the range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static IReadOnlyList<int> GetColumnNumbers(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var result = Enumerable.Range(range.FirstColumn + 1, range.ColumnCount).ToList();

            return result;
        }

        /// <summary>
        /// Gets the individual cells within a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The individual cells within the specified range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static IReadOnlyCollection<Cell> GetCells(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var result = new List<Cell>();

            var rowNumbers = range.GetRowNumbers();
            var columnNumbers = range.GetColumnNumbers();
            foreach (var rowNumber in rowNumbers)
            {
                foreach (var columnNumber in columnNumbers)
                {
                    var cell = range.Worksheet.GetCell(rowNumber, columnNumber);
                    result.Add(cell);
                }
            }

            return result;
        }

        /// <summary>
        /// Gets the individual cell ranges within a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The individual cell ranges within the specified range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static IReadOnlyCollection<Range> GetCellRanges(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var result = range.GetCells().Select(_ => _.GetRange()).ToList();

            return result;
        }

        /// <summary>
        /// Gets the cell area for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The cell area that covers the specified range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static CellArea GetCellArea(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var rowNumbers = range.GetRowNumbers();
            var columnNumbers = range.GetColumnNumbers();

            var result = new CellArea
            {
                StartRow = rowNumbers.First() - 1,
                EndRow = rowNumbers.Last() - 1,
                StartColumn = columnNumbers.First() - 1,
                EndColumn = columnNumbers.Last() - 1,
            };

            return result;
        }

        /// <summary>
        /// Gets the upper-left most cell in the specified range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The upper-left most cell in the specified range.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static Cell GetUpperLeftmostCell(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var result = range.Worksheet.Cells[range.FirstRow, range.FirstColumn];

            return result;
        }

        /// <summary>
        /// Gets the name of the range (e.g. A3:B5).
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// The name of the range (e.g. A3:B5).
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static string GetName(
            this Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            var rowNumbers = range.GetRowNumbers();
            var columnNumbers = range.GetColumnNumbers();

            string result;
            if ((rowNumbers.Count == 1) && (columnNumbers.Count == 1))
            {
                result = CellsHelper.CellIndexToName(rowNumbers.First() - 1, columnNumbers.First() - 1);
            }
            else
            {
                result = Invariant($"{CellsHelper.CellIndexToName(rowNumbers.First() - 1, columnNumbers.First() - 1)}:{CellsHelper.CellIndexToName(rowNumbers.Last() - 1, columnNumbers.Last() - 1)}");
            }

            return result;
        }
    }
}
