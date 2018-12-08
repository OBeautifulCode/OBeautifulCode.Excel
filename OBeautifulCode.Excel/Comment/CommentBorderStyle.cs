// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CommentBorderStyle.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    /// <summary>
    /// Specifies the style of a comment border.
    /// </summary>
    public enum CommentBorderStyle
    {
        /// <summary>
        /// Unknown (default).
        /// </summary>
        Unknown,

        /// <summary>
        /// Single line border.
        /// </summary>
        Single,

        /// <summary>
        /// Thick line between two thin lines.
        /// </summary>
        ThickBetweenThin,

        /// <summary>
        /// Thin line followed by a think line (outside-in).
        /// </summary>
        ThinThick,

        /// <summary>
        /// Thin line followed by a thin line (outside-in).
        /// </summary>
        ThickThin,

        /// <summary>
        /// Two thin lines.
        /// </summary>
        ThinThin,
    }
}
