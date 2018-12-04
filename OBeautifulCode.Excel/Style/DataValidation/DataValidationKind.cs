// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidationKind.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    /// <summary>
    /// Determines the kind of validation.
    /// </summary>
    public enum DataValidationKind
    {
        /// <summary>
        /// Any value is allowed.
        /// </summary>
        AnyValue,

        /// <summary>
        /// Only whole number are allowed.
        /// </summary>
        WholeNumber,

        /// <summary>
        /// Only decimals are allowed.
        /// </summary>
        Decimal,

        /// <summary>
        /// Only list-items are allowed.
        /// </summary>
        List,

        /// <summary>
        /// Only dates are allowed.
        /// </summary>
        Date,

        /// <summary>
        /// Only time is allowed.
        /// </summary>
        Time,

        /// <summary>
        /// Only text of a certain length is allowed.
        /// </summary>
        TextLength,

        /// <summary>
        /// Custom validation
        /// </summary>
        Custom,
    }
}
