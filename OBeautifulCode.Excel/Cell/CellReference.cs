﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReference.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Text.RegularExpressions;

    using OBeautifulCode.Assertion.Recipes;
    using OBeautifulCode.Type;

    using static System.FormattableString;

    /// <summary>
    /// Represents a reference to a cell.
    /// </summary>
    public partial class CellReference : IModelViaCodeGen, IDeclareToStringMethod
    {
        private static readonly Regex ValidWorksheetNameRegex = new Regex("^(?!.{32})(?=.*[\x20-\x26\x28-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E]$)[\x20-\x26\x28-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E][\x20-\x29\x2B-\x2E\x30-\x39\x3B-\x3E\x40-\x5A\x5E-\x7E]{0,30}$", RegexOptions.Compiled);

        private static readonly Regex ValidA1ReferenceRegex = new Regex(Invariant($"^[A-z]{{1,{Constants.MaximumColumnName.Length}}}[1-9][0-9]{{0,{Constants.MaximumRowNumber.ToString(CultureInfo.InvariantCulture).Length - 1}}}$"), RegexOptions.Compiled);

        private static readonly Regex ColumnNameInA1ReferenceRegex = new Regex("^[A-z]+", RegexOptions.Compiled);

        private static readonly Regex RowNumberInA1ReferenceRegex = new Regex("[0-9]+$", RegexOptions.Compiled);

        private static readonly CellReference KnownMissingCellReference = new CellReference(@" !""#$%&'()+,-.;<=>@^_`{|}~54320", 1, 1);

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
            new { worksheetName }.AsArg().Must().NotBeNullNorWhiteSpace().And().BeMatchedByRegex(ValidWorksheetNameRegex, "Worksheet names must have >= 1 and <= 31 characters.  The first or last character cannot be a single quote (').  Otherwise, all characters are allowed except for the characters in this set: {\\, /, *, [, ], :, ?}.");
            new { rowNumber }.AsArg().Must().BeGreaterThanOrEqualTo(1).And().BeLessThanOrEqualTo(Constants.MaximumRowNumber);
            new { columnNumber }.AsArg().Must().BeGreaterThanOrEqualTo(1).And().BeLessThanOrEqualTo(Constants.MaximumColumnNumber);

            this.WorksheetName = worksheetName;
            this.RowNumber = rowNumber;
            this.ColumnNumber = columnNumber;
        }

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string WorksheetName { get; private set; }

        /// <summary>
        /// Gets the 1-based row number.
        /// </summary>
        public int RowNumber { get; private set; }

        /// <summary>
        /// Gets the 1-based column number.
        /// </summary>
        public int ColumnNumber { get; private set; }

        /// <summary>
        /// Gets the reference to the cell, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        public string A1Reference => Invariant($"{CellsHelper.GetColumnName(this.ColumnNumber)}{this.RowNumber}");

        /// <summary>
        /// Gets the worksheet-qualified reference to the cell, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        public string WorksheetQualifiedA1Reference => Invariant($"'{this.WorksheetName.Replace("'", "''")}'!{this.A1Reference}");

        /// <summary>
        /// Gets a cell reference to a cell that is known (before de-referencing) to be missing.
        /// This is used when a cell reference object is required, but it is known/established that
        /// the cell is missing.
        /// </summary>
        /// <returns>
        /// A cell reference that indicates a known missing cell.
        /// </returns>
        public static CellReference GetKnownMissing()
        {
            var result = KnownMissingCellReference.DeepClone();

            return result;
        }

        /// <summary>
        /// Gets the <see cref="CellReference"/> equivalent to the specified reference to a cell in A1 notation (e.g. B4).
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="a1Reference">The cell reference in A1 notation.</param>
        /// <returns>
        /// The equivalent cell reference.
        /// </returns>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "a", Justification = "This is not hungarian notation.")]
        public static CellReference FromA1Reference(
            string worksheetName,
            string a1Reference)
        {
            new { a1Reference }.AsArg().Must().NotBeNullNorWhiteSpace().And().BeMatchedByRegex(ValidA1ReferenceRegex);

            var columnNameInReference = ColumnNameInA1ReferenceRegex.Match(a1Reference).Value;
            var rowNumberInReference = RowNumberInA1ReferenceRegex.Match(a1Reference).Value;

            var columnNumber = CellsHelper.GetColumnNumber(columnNameInReference);
            var rowNumber = int.Parse(rowNumberInReference, CultureInfo.InvariantCulture);

            var result = new CellReference(worksheetName, rowNumber, columnNumber);

            return result;
        }

        /// <summary>
        /// Gets the <see cref="CellReference"/> equivalent to the specified
        /// worksheet-qualified reference to a cell, using A1 notation (e.g. 'worksheet'!A5).
        /// </summary>
        /// <param name="worksheetQualifiedA1Reference">The worksheet-qualified reference to a cell, using A1 notation (e.g. 'worksheet'!A5).</param>
        /// <returns>
        /// The equivalent cell reference.
        /// </returns>
        public static CellReference FromWorksheetQualifiedA1Reference(
            string worksheetQualifiedA1Reference)
        {
            new { worksheetQualifiedA1Reference }.AsArg().Must().NotBeNullNorWhiteSpace().And().ContainString("!");

            var tokens = worksheetQualifiedA1Reference.Split(new[] { '!' }, 2);

            var worksheetNameToken = tokens[0];

            var worksheetNameTokenLength = worksheetNameToken.Length;
            new { worksheetNameTokenLength }.AsArg().Must().BeGreaterThanOrEqualTo(3);

            var worksheetNameTokenStartsWithApostrophe = worksheetNameToken.StartsWith("'", StringComparison.OrdinalIgnoreCase);
            new { worksheetNameTokenStartsWithApostrophe }.AsArg().Must().BeTrue();

            var worksheetNameTokenEndsWithApostrophe = worksheetNameToken.EndsWith("'", StringComparison.OrdinalIgnoreCase);
            new { worksheetNameTokenEndsWithApostrophe }.AsArg().Must().BeTrue();

            var worksheetName = worksheetNameToken.Remove(0, 1);
            worksheetName = worksheetName.Remove(worksheetName.Length - 1, 1);

            var a1ReferenceToken = tokens[1];

            var result = FromA1Reference(worksheetName, a1ReferenceToken);

            return result;
        }

        /// <summary>
        /// Determines if this object references a cell that is known (before de-referencing) to be missing.
        /// </summary>
        /// <returns>
        /// true if the cell is known to be missing; otherwise, false.
        /// </returns>
        public bool IsKnownMissing()
        {
            var result = this == KnownMissingCellReference;

            return result;
        }

        /// <inheritdoc cref="IDeclareToStringMethod.ToString" />
        public override string ToString() => this.IsKnownMissing() ? "KNOWN MISSING" : this.WorksheetQualifiedA1Reference;
    }
}
