// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellsHelper.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.String.Recipes;

    using static System.FormattableString;

    /// <summary>
    /// Helper methods related to cells.
    /// </summary>
    public static class CellsHelper
    {
        /// <summary>
        /// Gets the column name for the specified 1-based column number.
        /// </summary>
        /// <param name="columnNumber">The 1-based column number.</param>
        /// <returns>
        /// The column name.
        /// </returns>
        public static string GetColumnName(
            int columnNumber)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnNumber)}' < '{1}'"), (Exception)null);
            }

            if (columnNumber > Constants.MaximumColumnNumber)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnNumber)}' > '{Constants.MaximumColumnNumber}'"), (Exception)null);
            }

            var result = GetColumnNameInternal(columnNumber);

            return result;
        }

        /// <summary>
        /// Gets the 1-based column number for the specified column name.
        /// </summary>
        /// <param name="columnName">The column name.</param>
        /// <returns>
        /// The 1-based column number.
        /// </returns>
        public static int GetColumnNumber(
            string columnName)
        {
            if (columnName == null)
            {
                throw new ArgumentNullException(nameof(columnName));
            }

            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException(Invariant($"'{nameof(columnName)}' is white space"));
            }

            if (!columnName.IsAlphabetic())
            {
                throw new ArgumentException(Invariant($"'{nameof(columnName)}' is not alphabetic"));
            }

            var columnNameLength = columnName.Length;
            if (columnNameLength > Constants.MaximumColumnName.Length)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnNameLength)}' > '{Constants.MaximumColumnName.Length}'"), (Exception)null);
            }

            columnName = columnName.ToUpperInvariant();

            var columnNumber = 0;

            foreach (var columnNameCharacter in columnName)
            {
                columnNumber *= 26;
                columnNumber += columnNameCharacter - 'A' + 1;
            }

            if (columnNumber > Constants.MaximumColumnNumber)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnNumber)}' > '{Constants.MaximumColumnNumber}'"), (Exception)null);
            }

            var result = columnNumber;

            return result;
        }

        private static string GetColumnNameInternal(
            int columnNumber)
        {
            if (columnNumber < 1)
            {
                return string.Empty;
            }

            var result = GetColumnNameInternal((columnNumber - 1) / 26) + (char)('A' + ((columnNumber - 1) % 26));

            return result;
        }
    }
}
