// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CustomFormatString.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// Contains various custom format strings.
    /// </summary>
    public static class CustomFormatString
    {
        /// <summary>
        /// A date format that includes year, month and day (e.g. 2019-02-22).
        /// </summary>
        public static readonly string YearMonthDayDateFormat = "yyyy-mm-dd";

        /// <summary>
        /// A time format that includes hour and minute (e.g. 18:59).
        /// </summary>
        public static readonly string HourMinuteTimeFormat = "hh:mm";

        /// <summary>
        /// A number format for formatting a year (e.g. 1970).
        /// </summary>
        public static readonly string YearNumberFormat = "###0";

        /// <summary>
        /// A custom format that hides all values.
        /// </summary>
        public static readonly string HideCellValuesFormat = ";;;";

        /// <summary>
        /// Builds a number format using commas to separated thousands, and showing decimals
        /// the specified number of decimal places (e.g. 18,202.392).
        /// </summary>
        /// <param name="numberOfDecimalPlaces">The number of decimal places to show.</param>
        /// <returns>
        /// The number format for the specified number of decimal places.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="numberOfDecimalPlaces"/> is less than 1 or greater than 30.</exception>
        public static string BuildCommonSeparatedThousandsWithDecimalsNumberFormat(
            int numberOfDecimalPlaces)
        {
            new { numberOfDecimalPlaces }.Must().BeGreaterThanOrEqualTo(1);
            new { numberOfDecimalPlaces }.Must().BeLessThanOrEqualTo(30);

            var result = "#,##0." + new string('0', numberOfDecimalPlaces);
            return result;
        }
    }
}
