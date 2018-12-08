// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Format.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    /// <summary>
    /// Specifies all pre-canned formats.
    /// </summary>
    public enum Format
    {
        /// <summary>
        /// Unknown (default).
        /// </summary>
        Unknown,

        /// <summary>
        /// General format.
        /// </summary>
        General,

        /// <summary>
        /// Decimal1 format: <code>0</code>.
        /// </summary>
        Decimal1,

        /// <summary>
        /// Decimal2 format: <code>0.00</code>.
        /// </summary>
        Decimal2,

        /// <summary>
        /// Decimal3 format: <code>#,##0</code>.
        /// </summary>
        Decimal3,

        /// <summary>
        /// Decimal4 format: <code>#,##0.00</code>.
        /// </summary>
        Decimal4,

        /// <summary>
        /// Currency1 format: <code>$#,##0_);($#,##0)</code>.
        /// </summary>
        Currency1,

        /// <summary>
        /// Currency2 format: <code>$#,##0_);[Red]($#,##0)</code>.
        /// </summary>
        Currency2,

        /// <summary>
        /// Currency3 format: <code>$#,##0.00_);($#,##0.00)</code>.
        /// </summary>
        Currency3,

        /// <summary>
        /// Currency4 format: <code>$#,##0.00_);[Red]($#,##0.00)</code>.
        /// </summary>
        Currency4,

        /// <summary>
        /// Currency5 format: <code>_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)</code>.
        /// </summary>
        Currency5,

        /// <summary>
        /// Currency6 format: <code>_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)</code>.
        /// </summary>
        Currency6,

        /// <summary>
        /// Accounting1 format: <code>#,##0_);(#,##0)</code>.
        /// </summary>
        Accounting1,

        /// <summary>
        /// Accounting2 format: <code>#,##0_);[Red](#,##0)</code>.
        /// </summary>
        Accounting2,

        /// <summary>
        /// Accounting3 format: <code>#,##0.00_);(#,##0.00)</code>.
        /// </summary>
        Accounting3,

        /// <summary>
        /// Accounting4 format: <code>#,##0.00_);[Red](#,##0.00)</code>.
        /// </summary>
        Accounting4,

        /// <summary>
        /// Accounting5 format: <code>_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)</code>.
        /// </summary>
        Accounting5,

        /// <summary>
        /// Accounting6 format: <code>_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)</code>.
        /// </summary>
        Accounting6,

        /// <summary>
        /// Scientific1 format: <code>0.00E+00</code>.
        /// </summary>
        Scientific1,

        /// <summary>
        /// Scientific2 format: <code>##0.0E+0</code>.
        /// </summary>
        Scientific2,

        /// <summary>
        /// Percentage1 format: <code>0%</code>.
        /// </summary>
        Percentage1,

        /// <summary>
        /// Percentage2 format: <code>0.00%</code>.
        /// </summary>
        Percentage2,

        /// <summary>
        /// Fraction1 format: <code># ?/?</code>.
        /// </summary>
        Fraction1,

        /// <summary>
        /// Fraction2 format: <code># ??/??</code>.
        /// </summary>
        Fraction2,

        /// <summary>
        /// Date1 format: <code>m/d/yyyy</code>.
        /// </summary>
        Date1,

        /// <summary>
        /// Date2 format: <code>d-mmm-yy</code>.
        /// </summary>
        Date2,

        /// <summary>
        /// Date3 format: <code>d-mmm</code>.
        /// </summary>
        Date3,

        /// <summary>
        /// Date4 format: <code>mmm-yy</code>.
        /// </summary>
        Date4,

        /// <summary>
        /// Time1 format: <code>h:mm AM/PM</code>.
        /// </summary>
        Time1,

        /// <summary>
        /// Time2 format: <code>h:mm:ss AM/PM</code>.
        /// </summary>
        Time2,

        /// <summary>
        /// Time3 format: <code>h:mm</code>.
        /// </summary>
        Time3,

        /// <summary>
        /// Time4 format: <code>h:mm:ss</code>.
        /// </summary>
        Time4,

        /// <summary>
        /// Time5 format: <code>m/d/yyyy h:mm</code>.
        /// </summary>
        Time5,

        /// <summary>
        /// Time6 format: <code>mm:ss</code>.
        /// </summary>
        Time6,

        /// <summary>
        /// Time7 format: <code>[h]:mm:ss</code>.
        /// </summary>
        Time7,

        /// <summary>
        /// Time8 format: <code>mm:ss.0</code>.
        /// </summary>
        Time8,

        /// <summary>
        /// Text format: <code>@</code>.
        /// </summary>
        Text,
    }
}
