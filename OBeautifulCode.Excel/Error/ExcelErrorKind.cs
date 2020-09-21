// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelErrorKind.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.CodeAnalysis.Recipes;

    /// <summary>
    /// Specifies a kind of Excel error.
    /// </summary>
    /// <remarks>
    /// We tried to get use more descriptive names (e.g. DivideByZero) but Microsoft
    /// doesn't have any official documentation (or none we could find) for Excel errors
    /// and anyways we thought it would be more straightforward to just use the error notation
    /// which is familiar to Excel users.
    ///
    /// We set each enum value equal to what Excel returns when you call =ERROR.TYPE().
    /// </remarks>
    [SuppressMessage("Microsoft.Design", "CA1027:MarkEnumsWithFlags", Justification = ObcSuppressBecause.CA1027_MarkEnumsWithFlags_EnumValuesArePurposefullyNonContiguous)]
    public enum ExcelErrorKind
    {
        /// <summary>
        /// No error.
        /// </summary>
        None = 0,

        /// <summary>
        /// Error #NULL!
        /// </summary>
        Null = 1,

        /// <summary>
        /// Error #DIV/0!
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Div", Justification = ObcSuppressBecause.CA1704_IdentifiersShouldBeSpelledCorrectly_SpellingIsCorrectInContextOfTheDomain)]
        Div0 = 2,

        /// <summary>
        /// Error #VALUE!
        /// </summary>
        Value = 3,

        /// <summary>
        /// Error #REF!
        /// </summary>
        Ref = 4,

        /// <summary>
        /// Error #NAME?
        /// </summary>
        Name = 5,

        /// <summary>
        /// Error #NUM!
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Num", Justification = ObcSuppressBecause.CA1704_IdentifiersShouldBeSpelledCorrectly_SpellingIsCorrectInContextOfTheDomain)]
        Num = 6,

        /// <summary>
        /// Error  #N/A
        /// </summary>
        Na = 7,

        /// <summary>
        /// Error #GETTING_DATA
        /// </summary>
        GettingData = 8,

        /// <summary>
        /// Error #SPILL!
        /// </summary>
        Spill = 9,

        /// <summary>
        /// Error #UNKONWN!
        /// </summary>
        Unknown = 12,

        /// <summary>
        /// Error #FIELD!
        /// </summary>
        Field = 13,

        /// <summary>
        /// Error #CALC!
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Calc", Justification = ObcSuppressBecause.CA1704_IdentifiersShouldBeSpelledCorrectly_SpellingIsCorrectInContextOfTheDomain)]
        Calc = 14,
    }
}
