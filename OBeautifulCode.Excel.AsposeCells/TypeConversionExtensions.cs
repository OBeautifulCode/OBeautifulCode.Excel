// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TypeConversionExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using static System.FormattableString;

    /// <summary>
    /// Contains extension methods to convert to/from OBeautifulCode.Excel and Aspose.Cells types.
    /// </summary>
    public static class TypeConversionExtensions
    {
        /// <summary>
        /// Converts a <see cref="BorderEdges"/> to a <see cref="BorderType"/>.
        /// </summary>
        /// <param name="borderEdges">The border edges to convert.</param>
        /// <returns>
        /// A <see cref="BorderType"/> converted from a <see cref="BorderEdges"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="borderEdges"/> is <see cref="BorderEdges.Unknown"/>.</exception>
        public static BorderType ToBorderType(
            this BorderEdges borderEdges)
        {
            new { borderEdges }.Must().NotBeEqualTo(BorderEdges.Unknown);

            var result = default(BorderType);

            if (borderEdges.HasFlag(BorderEdges.BottomBorder))
            {
                result |= BorderType.BottomBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.DiagonalDown))
            {
                result |= BorderType.DiagonalDown;
            }

            if (borderEdges.HasFlag(BorderEdges.DiagonalUp))
            {
                result |= BorderType.DiagonalUp;
            }

            if (borderEdges.HasFlag(BorderEdges.LeftBorder))
            {
                result |= BorderType.LeftBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.RightBorder))
            {
                result |= BorderType.RightBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.TopBorder))
            {
                result |= BorderType.TopBorder;
            }

            return result;
        }

        /// <summary>
        /// Converts a <see cref="BorderStyle"/> to a <see cref="CellBorderType"/>.
        /// </summary>
        /// <param name="borderStyle">The border style to convert.</param>
        /// <returns>
        /// A <see cref="CellBorderType"/> converted from a <see cref="BorderStyle"/>.
        /// </returns>
        public static CellBorderType ToCellBorderType(
            this BorderStyle borderStyle)
        {
            switch (borderStyle)
            {
                case BorderStyle.DashDot:
                    return CellBorderType.DashDot;
                case BorderStyle.DashDotDot:
                    return CellBorderType.DashDotDot;
                case BorderStyle.Dashed:
                    return CellBorderType.Dashed;
                case BorderStyle.Dotted:
                    return CellBorderType.Dotted;
                case BorderStyle.Double:
                    return CellBorderType.Double;
                case BorderStyle.Hair:
                    return CellBorderType.Hair;
                case BorderStyle.Medium:
                    return CellBorderType.Medium;
                case BorderStyle.MediumDashDot:
                    return CellBorderType.MediumDashDot;
                case BorderStyle.MediumDashDotDot:
                    return CellBorderType.MediumDashDotDot;
                case BorderStyle.MediumDashed:
                    return CellBorderType.MediumDashed;
                case BorderStyle.None:
                    return CellBorderType.None;
                case BorderStyle.SlantedDashDot:
                    return CellBorderType.SlantedDashDot;
                case BorderStyle.Thick:
                    return CellBorderType.Thick;
                case BorderStyle.Thin:
                    return CellBorderType.Thin;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(BorderStyle)} is not supported: {borderStyle}"));
            }
        }
    }
}
