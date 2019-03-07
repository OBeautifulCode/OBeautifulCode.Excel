// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReference.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Text.RegularExpressions;

    using OBeautifulCode.Math.Recipes;
    using OBeautifulCode.Validation.Recipes;

    using static System.FormattableString;

    /// <summary>
    /// Represents a reference to a cell.
    /// </summary>
    public class CellReference : IEquatable<CellReference>
    {
        private static readonly Regex ValidWorksheetNameRegex = new Regex("^(?!.{32})(?=.*[\x20-\x26\x28-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E]$)[\x20-\x26\x28-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E][\x20-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E]{0,30}$", RegexOptions.Compiled);

        /// <summary>
        /// Initializes a new instance of the <see cref="CellReference"/> class.
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rowNumber">The 1-based row number.</param>
        /// <param name="columnNumber">The 1-based column number.</param>
        public CellReference(
            string worksheetName,
            int rowNumber,
            int columnNumber)
        {
            new { worksheetName }.Must().NotBeNullNorWhiteSpace().And().BeMatchedByRegex(ValidWorksheetNameRegex, "Worksheet names must have >= 1 and <= 31 characters.  The first or last character cannot be a single quote (').  Otherwise, all characters are allowed except for the characters in this set: {\\, /, *, [, ], :, ?}.");
            new { rowNumber }.Must().BeGreaterThanOrEqualTo(1).And().BeLessThanOrEqualTo(Constants.MaximumRowNumber);
            new { columnNumber }.Must().BeGreaterThanOrEqualTo(1).And().BeLessThanOrEqualTo(Constants.MaximumColumnNumber);

            this.WorksheetName = worksheetName;
            this.RowNumber = rowNumber;
            this.ColumnNumber = columnNumber;
        }

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
        public string WorksheetName { get; private set; }

        /// <summary>
        /// Gets the 1-based row number.
        /// </summary>
        // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
        public int RowNumber { get; private set; }

        /// <summary>
        /// Gets the 1-based column number.
        /// </summary>
        // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
        public int ColumnNumber { get; private set; }

        /// <summary>
        /// Gets the worksheet-qualified reference, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        public string WorksheetQualifiedA1Reference => Invariant($"'{this.WorksheetName.Replace("'", "''")}'!{CellsHelper.GetColumnName(this.ColumnNumber)}{this.RowNumber}");

        /// <summary>
        /// Determines whether two objects of type <see cref="CellReference"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            CellReference item1,
            CellReference item2)
        {
            if (ReferenceEquals(item1, item2))
            {
                return true;
            }

            if (ReferenceEquals(item1, null) || ReferenceEquals(item2, null))
            {
                return false;
            }

            var result =
                (item1.WorksheetName == item2.WorksheetName) &&
                (item1.RowNumber == item2.RowNumber) &&
                (item1.ColumnNumber == item2.ColumnNumber);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellReference"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            CellReference item1,
            CellReference item2)
            => !(item1 == item2);

        /// <inheritdoc />
        public bool Equals(CellReference other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as CellReference);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.WorksheetName)
                .Hash(this.RowNumber)
                .Hash(this.ColumnNumber)
                .Value;
    }
}
