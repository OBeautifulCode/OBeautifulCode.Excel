// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelErrorKindExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    /// <summary>
    /// Extensions on <see cref="ExcelErrorKind"/>.
    /// </summary>
    public static class ExcelErrorKindExtensions
    {
        /// <summary>
        /// Gets the Excel identifier for the specified kind of error.
        /// </summary>
        /// <param name="errorKind">The kind of error.</param>
        /// <returns>
        /// The Excel identifier for the specified kind of error.
        /// </returns>
        public static string ToExcelIdentifier(
            this ExcelErrorKind errorKind)
        {
            switch (errorKind)
            {
                case ExcelErrorKind.None:
                    return null;
                case ExcelErrorKind.Null:
                    return "#NULL!";
                case ExcelErrorKind.Div0:
                    return "#DIV/0!";
                case ExcelErrorKind.Value:
                    return "#VALUE!";
                case ExcelErrorKind.Ref:
                    return "#REF!";
                case ExcelErrorKind.Name:
                    return "#NAME?";
                case ExcelErrorKind.Num:
                    return "#NUM!";
                case ExcelErrorKind.Na:
                    return "#N/A";
                case ExcelErrorKind.GettingData:
                    return "#GETTING_DATA";
                case ExcelErrorKind.Spill:
                    return "#SPILL!";
                case ExcelErrorKind.Unknown:
                    return "#UNKNOWN!";
                case ExcelErrorKind.Field:
                    return "#FIELD!";
                case ExcelErrorKind.Calc:
                    return "#CALC!";
                default:
                    throw new NotSupportedException("This error kind is not supported: " + errorKind);
            }
        }
    }
}
