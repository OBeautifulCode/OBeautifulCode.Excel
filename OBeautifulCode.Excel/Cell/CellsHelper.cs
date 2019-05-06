// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellsHelper.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Validation.Recipes;

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
            new { columnNumber }.Must().BeGreaterThanOrEqualTo(1).And().BeLessThanOrEqualTo(Constants.MaximumColumnNumber);

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
            new { columnName }.Must().NotBeNullNorWhiteSpace().And().BeAlphabetic();
            var columnNameLength = columnName.Length;
            new { columnNameLength }.Must().BeLessThanOrEqualTo(Constants.MaximumColumnName.Length);

            columnName = columnName.ToUpperInvariant();

            var columnNumber = 0;

            foreach (var columnNameCharacter in columnName)
            {
                columnNumber *= 26;
                columnNumber += columnNameCharacter - 'A' + 1;
            }

            new { columnNumber }.Must().BeLessThanOrEqualTo(Constants.MaximumColumnNumber);

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
