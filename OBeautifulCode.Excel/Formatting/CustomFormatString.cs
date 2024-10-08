// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CustomFormatString.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using OBeautifulCode.CodeAnalysis.Recipes;
    using OBeautifulCode.Type;
    using static System.FormattableString;

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
            if (numberOfDecimalPlaces < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(numberOfDecimalPlaces)}' < '{1}'"), (Exception)null);
            }

            if (numberOfDecimalPlaces > 30)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(numberOfDecimalPlaces)}' > '{30}'"), (Exception)null);
            }

            var result = "#,##0." + new string('0', numberOfDecimalPlaces);
            return result;
        }

        /// <summary>
        /// Converts a <see cref="DateTimeFormatKind"/> to it's equivalent Excel custom format string.
        /// </summary>
        /// <param name="dateTimeFormatKind">The format kind to use.</param>
        /// <param name="cultureKind">The culture kind to use.</param>
        /// <returns>
        /// The Excel custom format string that's equivalent to the the specified <see cref="DateTimeFormatKind"/>.
        /// </returns>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = ObcSuppressBecause.CA1502_AvoidExcessiveComplexity_DisagreeWithAssessment)]
        public static string ToExcelCustomFormatString(
            this DateTimeFormatKind dateTimeFormatKind,
            CultureKind cultureKind = CultureKind.Invariant)
        {
            if (dateTimeFormatKind == DateTimeFormatKind.Unknown)
            {
                throw new ArgumentException(Invariant($"{nameof(dateTimeFormatKind)} is {nameof(DateTimeFormatKind)}.{nameof(DateTimeFormatKind.Unknown)}."));
            }

            if (cultureKind == CultureKind.Unknown)
            {
                throw new ArgumentException(Invariant($"{nameof(cultureKind)} is {nameof(CultureKind)}.{nameof(CultureKind.Unknown)}."));
            }

            string result;

            if (cultureKind == CultureKind.Invariant)
            {
                switch (dateTimeFormatKind)
                {
                    case DateTimeFormatKind.ShortDatePattern:
                        result = "mm/dd/yyyy";
                        break;
                    case DateTimeFormatKind.LongDatePattern:
                        result = "dddd, dd mmmm yyyy";
                        break;
                    case DateTimeFormatKind.FullDateTimePatternShortTime:
                        result = "dddd, dd mmmm yyyy hh:mm";
                        break;
                    case DateTimeFormatKind.FullDateTimePatternLongTime:
                        result = "dddd, dd mmmm yyyy hh:mm:ss";
                        break;
                    case DateTimeFormatKind.GeneralDateTimePatternShortTime:
                        result = "mm/dd/yyyy hh:mm";
                        break;
                    case DateTimeFormatKind.GeneralDateTimePatternLongTime:
                        result = "mm/dd/yyyy hh:mm:ss";
                        break;
                    case DateTimeFormatKind.MonthDayPattern:
                        result = "mmmm dd";
                        break;
                    case DateTimeFormatKind.SortableDateTimePattern:
                        result = "yyyy-mm-dd\"T\"hh:mm:ss";
                        break;
                    case DateTimeFormatKind.ShortTimePattern:
                        result = "hh:mm";
                        break;
                    case DateTimeFormatKind.LongTimePattern:
                        result = "hh:mm:ss";
                        break;
                    case DateTimeFormatKind.UniversalSortableDateTimePattern:
                        result = "yyyy-mm-dd hh:mm:ss\"Z\"";
                        break;
                    case DateTimeFormatKind.UniversalFullDateTimePattern:
                        result = "dddd, dd mmmm yyyy hh:mm:ss";
                        break;
                    case DateTimeFormatKind.YearMonthPattern:
                        result = "yyyy mmmm";
                        break;
                    default:
                        throw new NotSupportedException(Invariant($"This {nameof(DateTimeFormatKind)} is not supported: {dateTimeFormatKind}."));
                }
            }
            else if (cultureKind == CultureKind.EnglishUnitedStates)
            {
                switch (dateTimeFormatKind)
                {
                    case DateTimeFormatKind.ShortDatePattern:
                        result = "m/d/yyyy";
                        break;
                    case DateTimeFormatKind.LongDatePattern:
                        result = "dddd, mmmm d, yyyy";
                        break;
                    case DateTimeFormatKind.FullDateTimePatternShortTime:
                        result = "dddd, mmmm d, yyyy h:mm AM/PM";
                        break;
                    case DateTimeFormatKind.FullDateTimePatternLongTime:
                        result = "dddd, mmmm d, yyyy h:mm:ss AM/PM";
                        break;
                    case DateTimeFormatKind.GeneralDateTimePatternShortTime:
                        result = "m/d/yyyy h:mm AM/PM";
                        break;
                    case DateTimeFormatKind.GeneralDateTimePatternLongTime:
                        result = "m/d/yyyy h:mm:ss AM/PM";
                        break;
                    case DateTimeFormatKind.MonthDayPattern:
                        result = "mmmm d";
                        break;
                    case DateTimeFormatKind.SortableDateTimePattern:
                        result = "yyyy-mm-dd\"T\"hh:mm:ss";
                        break;
                    case DateTimeFormatKind.ShortTimePattern:
                        result = "h:mm AM/PM";
                        break;
                    case DateTimeFormatKind.LongTimePattern:
                        result = "h:mm:ss AM/PM";
                        break;
                    case DateTimeFormatKind.UniversalSortableDateTimePattern:
                        result = "yyyy-mm-dd hh:mm:ss\"Z\"";
                        break;
                    case DateTimeFormatKind.UniversalFullDateTimePattern:
                        result = "dddd, mmmm d, yyyy h:mm:ss AM/PM";
                        break;
                    case DateTimeFormatKind.YearMonthPattern:
                        result = "mmmm yyyy";
                        break;
                    default:
                        throw new NotSupportedException(Invariant($"This {nameof(DateTimeFormatKind)} is not supported: {dateTimeFormatKind}."));
                }
            }
            else
            {
                throw new NotImplementedException(Invariant($"This {nameof(CultureKind)} is not yet implemented: {cultureKind}."));
            }

            return result;
        }
    }
}
