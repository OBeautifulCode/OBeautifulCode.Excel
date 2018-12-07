// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StyleContainer.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Container for style-related objects.
    /// </summary>
    public class StyleContainer
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StyleContainer"/> class.
        /// </summary>
        /// <param name="style">The style.</param>
        /// <param name="styleFlag">The style flag.</param>
        /// <exception cref="ArgumentNullException"><paramref name="style"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="styleFlag"/> is null.</exception>
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flag", Justification = "This is the best word in the parameter name.")]
        public StyleContainer(
            Style style,
            StyleFlag styleFlag)
        {
            new { style }.Must().NotBeNull();
            new { styleFlag }.Must().NotBeNull();

            this.Style = style;
            this.StyleFlag = styleFlag;
        }

        /// <summary>
        /// Gets the style.
        /// </summary>
        public Style Style { get; }

        /// <summary>
        /// Gets the style flag.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flag", Justification = "This is the best word in the property name.")]
        public StyleFlag StyleFlag { get; }

        /// <summary>
        /// Builds a style-container using a new style for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        /// A style container.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static StyleContainer BuildNewFromRange(
            Range range)
        {
            new { range }.Must().NotBeNull();

            var style = range.Worksheet.Workbook.CreateStyle();
            var styleFlag = new StyleFlag();

            var result = new StyleContainer(style, styleFlag);

            return result;
        }

        /// <summary>
        /// Builds a style-container using the existing style for the cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// A style container.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static StyleContainer BuildExistingFromCell(
            Cell cell)
        {
            new { cell }.Must().NotBeNull();

            var style = cell.GetStyle();
            var styleFlag = new StyleFlag();

            var result = new StyleContainer(style, styleFlag);

            return result;
        }

        /// <summary>
        /// Applies this style container to the specified range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public void ApplyToRange(
            Range range)
        {
            new { range }.Must().NotBeNull();

            range.ApplyStyle(this.Style, this.StyleFlag);
        }

        /// <summary>
        /// Applies this style container to the specified cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public void ApplyToCell(
            Cell cell)
        {
            new { cell }.Must().NotBeNull();

            cell.SetStyle(this.Style, this.StyleFlag);
        }
    }
}
