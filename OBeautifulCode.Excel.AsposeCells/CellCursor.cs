// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellCursor.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using Aspose.Cells;

    using OBeautifulCode.Assertion.Recipes;

    using static System.FormattableString;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Provides methods for navigating a virtual spreadsheet, tracking the space visited (the "canvassed" space),
    /// and marking/tagging cells for quick access.
    /// </summary>
    /// <remarks>
    /// The "canvassed" space is considered to be between the starting row and column and the max row and column
    /// visited (independently).  Not all cells within that space need to be explicitly visited to be considered
    /// canvassed.  For example, if we start at A1, move right by 3 and then down by 3, we will have never
    /// visited A2, but A2 is considered to have been canvassed.
    /// </remarks>
    public class CellCursor
    {
        private IDictionary<string, List<Cell>> markerNameToCellsMap = new Dictionary<string, List<Cell>>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CellCursor"/> class.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowNumber">The starting/current row number.</param>
        /// <param name="columnNumber">The starting/current column number.</param>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="rowNumber"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnNumber"/> is less than 1.</exception>
        public CellCursor(
            Worksheet worksheet,
            int rowNumber = 1,
            int columnNumber = 1)
        {
            new { worksheet }.AsArg().Must().NotBeNull();
            new { rowNumber }.AsArg().Must().BeGreaterThanOrEqualTo(1);
            new { columnNumber }.AsArg().Must().BeGreaterThanOrEqualTo(1);

            this.Worksheet = worksheet;
            this.RowNumber = rowNumber;
            this.ColumnNumber = columnNumber;

            this.MaxRowNumber = rowNumber;
            this.MaxColumnNumber = columnNumber;

            this.StartRowNumber = rowNumber;
            this.StartColumnNumber = columnNumber;
        }

        /// <summary>
        /// Gets the worksheet.
        /// </summary>
        public Worksheet Worksheet { get; }

        /// <summary>
        /// Gets the current row.
        /// </summary>
        public int RowNumber { get; private set; }

        /// <summary>
        /// Gets the current column number.
        /// </summary>
        public int ColumnNumber { get; private set; }

        /// <summary>
        /// Gets the starting row number.
        /// </summary>
        public int StartRowNumber { get; private set; }

        /// <summary>
        /// Gets the starting column number.
        /// </summary>
        public int StartColumnNumber { get; private set; }

        /// <summary>
        /// Gets the highest row number the cursor has been on.
        /// </summary>
        public int MaxRowNumber { get; private set; }

        /// <summary>
        /// Gets the highest column number the cursor has been on.
        /// </summary>
        public int MaxColumnNumber { get; private set; }

        /// <summary>
        /// Gets the cell at the cursor.
        /// </summary>
        public Cell Cell => this.Worksheet.GetCell(this.RowNumber, this.ColumnNumber);

        /// <summary>
        /// Gets the range of the cell at the cursor.
        /// </summary>
        public Range CellRange => this.Worksheet.GetRange(this.RowNumber, this.RowNumber, this.ColumnNumber, this.ColumnNumber);

        /// <summary>
        /// Gets the range of cells that have been canvassed for the row at the cursor.
        /// </summary>
        public Range CanvassedRowRange => this.Worksheet.GetRange(this.RowNumber, this.RowNumber, this.StartColumnNumber, this.MaxColumnNumber);

        /// <summary>
        /// Gets the range of cells that have been canvassed for the column at the cursor.
        /// </summary>
        public Range CanvassedColumnRange => this.Worksheet.GetRange(this.StartRowNumber, this.MaxRowNumber, this.ColumnNumber, this.ColumnNumber);

        /// <summary>
        /// Gets the range of cells that have been canvassed.
        /// </summary>
        public Range CanvassedRange => this.Worksheet.GetRange(this.StartRowNumber, this.MaxRowNumber, this.StartColumnNumber, this.MaxColumnNumber);

        /// <summary>
        /// Clones this cell cursor, but with a specified worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns>
        /// A clone of this cell cursor, except with the worksheet swapped-out for the specified worksheet.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        public CellCursor CloneWithWorksheet(
            Worksheet worksheet)
        {
            new { worksheet }.AsArg().Must().NotBeNull();

            var result = new CellCursor(worksheet)
            {
                RowNumber = this.RowNumber,
                ColumnNumber = this.ColumnNumber,
                StartRowNumber = this.StartRowNumber,
                StartColumnNumber = this.StartColumnNumber,
                MaxRowNumber = this.MaxRowNumber,
                MaxColumnNumber = this.MaxColumnNumber,
                markerNameToCellsMap = this.markerNameToCellsMap.ToDictionary(
                    _ => _.Key,
                    _ => _.Value.Select(cell => worksheet.Cells[cell.Row, cell.Column]).ToList()),
            };

            return result;
        }

        /// <summary>
        /// Resets the cursor to its original state.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor Reset()
        {
            this.ResetRow();
            this.ResetColumn();

            return this;
        }

        /// <summary>
        /// Resets the row of the cursor to its original state.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor ResetRow()
        {
            this.RowNumber = this.StartRowNumber;

            return this;
        }

        /// <summary>
        /// Resets the column of the cursor to its original state.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor ResetColumn()
        {
            this.ColumnNumber = this.StartColumnNumber;

            return this;
        }

        /// <summary>
        /// Moves to the cursor to the bottom-right of the canvas.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor MoveToBottomRightOfCanvas()
        {
            this.MoveRightToMaxColumn();
            this.MoveDownToMaxRow();

            return this;
        }

        /// <summary>
        /// Moves to the cursor all the way to the right
        /// of the canvas, maintaining the current row.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor MoveRightToMaxColumn()
        {
            this.ColumnNumber = this.MaxColumnNumber;

            return this;
        }

        /// <summary>
        /// Moves to the cursor all the way to the bottom
        /// of the canvas, maintaining the current column.
        /// </summary>
        /// <returns>
        /// This cursor.
        /// </returns>
        public CellCursor MoveDownToMaxRow()
        {
            this.RowNumber = this.MaxRowNumber;

            return this;
        }

        /// <summary>
        /// Moves the cursor down by a specified number of cells.
        /// </summary>
        /// <param name="by">The number of cells to move down by.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="by"/> is less than 0.</exception>
        public CellCursor MoveDown(
            int by = 1)
        {
            new { by }.AsArg().Must().BeGreaterThanOrEqualTo(0);

            this.RowNumber = this.RowNumber + by;
            this.MaxRowNumber = this.RowNumber > this.MaxRowNumber ? this.RowNumber : this.MaxRowNumber;

            return this;
        }

        /// <summary>
        /// Moves the cursor up by a specified number of cells.
        /// </summary>
        /// <param name="by">The number of cells to move up by.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="by"/> is less than 0.</exception>
        /// <exception cref="InvalidOperationException">Moving up by <paramref name="by"/> places the cursor above <see cref="StartRowNumber"/>.</exception>
        public CellCursor MoveUp(
            int by = 1)
        {
            new { by }.AsArg().Must().BeGreaterThanOrEqualTo(0);

            if (this.RowNumber - by < this.StartRowNumber)
            {
                throw new InvalidOperationException(Invariant($"Cannot move up by {by} because it would place cursor above start row {this.StartRowNumber}"));
            }

            this.RowNumber = this.RowNumber - by;

            return this;
        }

        /// <summary>
        /// Moves the cursor right by a specified number of cells.
        /// </summary>
        /// <param name="by">The number of cells to move down by.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="by"/> is less than 0.</exception>
        public CellCursor MoveRight(
            int by = 1)
        {
            new { by }.AsArg().Must().BeGreaterThanOrEqualTo(0);

            this.ColumnNumber = this.ColumnNumber + by;
            this.MaxColumnNumber = this.ColumnNumber > this.MaxColumnNumber ? this.ColumnNumber : this.MaxColumnNumber;

            return this;
        }

        /// <summary>
        /// Moves the cursor left by a specified number of cells.
        /// </summary>
        /// <param name="by">The number of cells to move left by.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="by"/> is less than 0.</exception>
        /// <exception cref="InvalidOperationException">Moving left by <paramref name="by"/> places the cursor to the left of <see cref="StartColumnNumber"/>.</exception>
        public CellCursor MoveLeft(
            int by = 1)
        {
            new { by }.AsArg().Must().BeGreaterThanOrEqualTo(0);

            if (this.ColumnNumber - by < this.StartColumnNumber)
            {
                throw new InvalidOperationException(Invariant($"Cannot move left by {by} because it would place cursor to the left of start column {this.StartColumnNumber}"));
            }

            this.ColumnNumber = this.ColumnNumber - by;

            return this;
        }

        /// <summary>
        /// Adds a marker to the current cell.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// This cursor.  If the specified marker doesn't already exist for the current cell, then it will be added.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        public CellCursor AddMarker(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            if (!this.markerNameToCellsMap.ContainsKey(markerName))
            {
                this.markerNameToCellsMap.Add(markerName, new List<Cell>());
            }

            var cells = this.markerNameToCellsMap[markerName];
            if (!cells.Select(_ => _.Name).Contains(this.Cell.Name))
            {
                cells.Add(this.Cell);
            }

            return this;
        }

        /// <summary>
        /// Removes all marker.
        /// </summary>
        public void RemoveAllMarkers()
        {
            this.markerNameToCellsMap = new Dictionary<string, List<Cell>>();
        }

        /// <summary>
        /// Removes a marker.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        public CellCursor RemoveMarker(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            if (this.markerNameToCellsMap.ContainsKey(markerName))
            {
                this.markerNameToCellsMap.Remove(markerName);
            }

            return this;
        }

        /// <summary>
        /// Determines if the specified marker has been applied.
        /// </summary>
        /// <param name="markerName">The case-sensitive marker name.</param>
        /// <returns>
        /// true if the marker has been applied; otherwise, false.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        public bool HasMarker(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            var result = this.markerNameToCellsMap.ContainsKey(markerName);

            return result;
        }

        /// <summary>
        /// Gets the marked range.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// The range of the marker.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        /// <exception cref="InvalidOperationException">The marker doesn't exist.</exception>
        public Range GetMarkedRange(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            if (!this.HasMarker(markerName))
            {
                throw new InvalidOperationException("marker does not exist: " + markerName);
            }

            var cells = this.markerNameToCellsMap[markerName];

            var firstRow = cells.Min(_ => _.GetRowNumber());
            var lastRow = cells.Max(_ => _.GetRowNumber());
            var firstColumn = cells.Min(_ => _.GetColumnNumber());
            var lastColumn = cells.Max(_ => _.GetColumnNumber());

            var result = this.Worksheet.GetRange(firstRow, lastRow, firstColumn, lastColumn);

            return result;
        }

        /// <summary>
        /// Gets the marked cells, in the order they were marked.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// The marked cells, in the order they were marked.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        /// <exception cref="InvalidOperationException">The marker does not exist.</exception>
        public IReadOnlyList<Cell> GetMarkedCells(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            if (!this.HasMarker(markerName))
            {
                throw new InvalidOperationException("marker does not exist: " + markerName);
            }

            var result = this.markerNameToCellsMap[markerName].ToList();

            return result;
        }

        /// <summary>
        /// Gets the marked cell.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// The marked cell.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        /// <exception cref="InvalidOperationException">The marker does not exist.</exception>
        /// <exception cref="InvalidOperationException">Multiple cells are marked with <paramref name="markerName"/>.</exception>
        public Cell GetMarkedCell(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            var markedCells = this.GetMarkedCells(markerName);
            if (markedCells.Count > 1)
            {
                throw new InvalidOperationException("Multiple cells are marked as: " + markerName);
            }

            var result = markedCells.SingleOrDefault();

            if (result == null)
            {
                throw new InvalidOperationException("Something went wrong: marker exists but no cells are marked as: " + markerName);
            }

            return result;
        }

        /// <summary>
        /// Gets a reference to the marked cell.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// A reference to the marked cell.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        /// <exception cref="InvalidOperationException">The marker does not exist.</exception>
        /// <exception cref="InvalidOperationException">Multiple cells are marked with <paramref name="markerName"/>.</exception>
        public CellReference GetMarkedCellReference(
            string markerName)
        {
            var cell = this.GetMarkedCell(markerName);

            var result = cell.ToCellReference();

            return result;
        }

        /// <summary>
        /// Moves the cursor to the marked cell.
        /// </summary>
        /// <param name="markerName">The case-sensitive name of the marker.</param>
        /// <returns>
        /// This cursor.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="markerName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="markerName"/> is white space.</exception>
        /// <exception cref="InvalidOperationException">The marker does not exist.</exception>
        /// <exception cref="InvalidOperationException">Multiple cells are marked with <paramref name="markerName"/>.</exception>
        public CellCursor MoveToMarkedCell(
            string markerName)
        {
            new { markerName }.AsArg().Must().NotBeNullNorWhiteSpace();

            var markedCell = this.GetMarkedCell(markerName);

            this.RowNumber = markedCell.GetRowNumber();
            this.ColumnNumber = markedCell.GetColumnNumber();

            return this;
        }
    }
}
