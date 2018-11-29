// --------------------------------------------------------------------------------------------------------------------
// <copyright file="BorderEdges.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    /// <summary>
    /// Specifies the edges of the border to apply styling to.
    /// </summary>
    /// <remarks>
    /// This mirrors Aspose.Cells.BorderType.
    /// </remarks>
    [Flags]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue", Justification = "Unknown is the right term here.  There is no notion of a 'None' border edge.")]
    public enum BorderEdges
    {
        /// <summary>
        /// Unknown (default).
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Left border.
        /// </summary>
        LeftBorder = 1,

        /// <summary>
        /// Right border.
        /// </summary>
        RightBorder = 2,

        /// <summary>
        /// Top border.
        /// </summary>
        TopBorder = 4,

        /// <summary>
        /// Bottom border.
        /// </summary>
        BottomBorder = 8,

        /// <summary>
        /// Diagonal down border.
        /// </summary>
        DiagonalDown = 16,

        /// <summary>
        /// Diagonal up border.
        /// </summary>
        DiagonalUp = 32,

        /// <summary>
        /// Outline border - left + right + top + bottom
        /// </summary>
        Outline = LeftBorder | RightBorder | TopBorder | BottomBorder,
    }
}
