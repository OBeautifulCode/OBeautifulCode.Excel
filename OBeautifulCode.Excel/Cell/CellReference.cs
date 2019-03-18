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
    using OBeautifulCode.Type;
    using OBeautifulCode.Validation.Recipes;

    using static System.FormattableString;

    /// <summary>
    /// Represents a reference to a cell.
    /// </summary>
    public class CellReference : IEquatable<CellReference>, IDeepCloneable<CellReference>
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
        /// Gets the reference to the cell, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        public string A1Reference => Invariant($"{CellsHelper.GetColumnName(this.ColumnNumber)}{this.RowNumber}");

        /// <summary>
        /// Gets the worksheet-qualified reference to the cell, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        public string WorksheetQualifiedA1Reference => Invariant($"'{this.WorksheetName.Replace("'", "''")}'!{A1Reference}");

        /// <summary>
        /// Determines whether two objects of type <see cref="CellReference"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            CellReference left,
            CellReference right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
            {
                return false;
            }

            var result =
                (left.WorksheetName == right.WorksheetName) &&
                (left.RowNumber == right.RowNumber) &&
                (left.ColumnNumber == right.ColumnNumber);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellReference"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            CellReference left,
            CellReference right)
            => !(left == right);

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

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public CellReference DeepClone()
        {
            var result = new CellReference(this.WorksheetName, this.RowNumber, this.ColumnNumber);

            return result;
        }
    }
}
