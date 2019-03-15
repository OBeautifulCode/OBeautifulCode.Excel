// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeStyleExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Drawing;

    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// Extension methods on type <see cref="RangeStyle"/>.
    /// </summary>
    public static class RangeStyleExtensions
    {
        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with the specified font color set.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="fontColor">The font color to set.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified font color set.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="rangeStyle"/> is null.</exception>
        public static RangeStyle DeepCloneWithFontColor(
            this RangeStyle rangeStyle,
            Color? fontColor)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.FontColor = fontColor;

            return result;
        }

        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with the specified font size set.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="fontSize">The font size to set.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified font size set.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="rangeStyle"/> is null.</exception>
        public static RangeStyle DeepCloneWithFontSize(
            this RangeStyle rangeStyle,
            int? fontSize)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.FontSize = fontSize;

            return result;
        }

        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with a specified value indicating whether the font should be bold or not.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="fontIsBold">True to bold the font; otherwise, false.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified value indicating whether the font should be bold or not.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="rangeStyle"/> is null.</exception>
        public static RangeStyle DeepCloneWithFontIsBold(
            this RangeStyle rangeStyle,
            bool? fontIsBold)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.FontIsBold = fontIsBold;

            return result;
        }

        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with the specified background color set.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="backgroundColor">The background color to set.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified background color set.
        /// </returns>
        public static RangeStyle DeepCloneWithBackgroundColor(
            this RangeStyle rangeStyle,
            Color? backgroundColor)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.BackgroundColor = backgroundColor;

            return result;
        }

        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with the specified row height set.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="rowHeightInPixels">The row height, in pixels.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified row height set.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="rangeStyle"/> is null.</exception>
        public static RangeStyle DeepCloneWithRowHeightInPixels(
            this RangeStyle rangeStyle,
            int? rowHeightInPixels)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.RowHeightInPixels = rowHeightInPixels;

            return result;
        }

        /// <summary>
        /// Deep clones the specified <see cref="RangeStyle"/>, but with the specified column width set.
        /// </summary>
        /// <param name="rangeStyle">The range style to clone.</param>
        /// <param name="columnWidthInPixels">The column width, in pixels.</param>
        /// <returns>
        /// A clone of the specified range style, but with the specified column width set.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="rangeStyle"/> is null.</exception>
        public static RangeStyle DeepCloneWithColumnWidthInPixels(
            this RangeStyle rangeStyle,
            int? columnWidthInPixels)
        {
            new { rangeStyle }.Must().NotBeNull();

            var result = rangeStyle.DeepClone();
            result.ColumnWidthInPixels = columnWidthInPixels;

            return result;
        }
    }
}
