// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidationOperator.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Specifies the operator to use to validate the data.
    /// </summary>
    public enum DataValidationOperator
    {
        /// <summary>
        /// Data must be between two values.
        /// </summary>
        Between,

        /// <summary>
        /// Data must be equal to some value.
        /// </summary>
        EqualTo,

        /// <summary>
        /// Data must be greater than some value.
        /// </summary>
        GreaterThan,

        /// <summary>
        /// Data must be greater than or equal to some value.
        /// </summary>
        GreaterThanOrEqualTo,

        /// <summary>
        /// Data must be less than some value.
        /// </summary>
        LessThan,

        /// <summary>
        /// Data must be less than or equal to some value.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LessOr", Justification = "This is not a compound word.")]
        LessThanOrEqualTo,

        /// <summary>
        /// No operator is applicable.
        /// </summary>
        None,

        /// <summary>
        /// Data must not be between two values.
        /// </summary>
        NotBetween,

        /// <summary>
        /// Data must not be equal to some value.
        /// </summary>
        NotEqualTo,
    }
}
