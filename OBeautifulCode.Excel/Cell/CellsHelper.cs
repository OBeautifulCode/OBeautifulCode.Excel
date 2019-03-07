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
