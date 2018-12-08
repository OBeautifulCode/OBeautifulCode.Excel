// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetExtensions.Write.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Drawing;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// Extensions methods on type <see cref="Worksheet"/>.
    /// </summary>
    public static partial class WorksheetExtensions
    {
        /// <summary>
        /// Sets the worksheet tab color.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="color">The color.</param>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        public static void SetTabColor(
            this Worksheet worksheet,
            Color? color)
        {
            new { worksheet }.Must().NotBeNull();

            if (color != null)
            {
                worksheet.TabColor = (Color)color;
            }
        }

        /// <summary>
        /// Sets the worksheet visibility.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="isHidden">Determines if the worksheet should be hidden or not (visible).</param>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        public static void SetWorksheetVisibility(
            this Worksheet worksheet,
            bool? isHidden)
        {
            new { worksheet }.Must().NotBeNull();

            if (isHidden != null)
            {
                worksheet.IsVisible = !(bool)isHidden;
            }
        }

        /// <summary>
        /// Sets the row and column headings visibility.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="isHidden">Determines if the row and column headings are hidden or not (visible).</param>
        /// <exception cref="ArgumentNullException"><paramref name="worksheet"/> is null.</exception>
        public static void SetRowAndColumnHeadingsVisibility(
            this Worksheet worksheet,
            bool? isHidden)
        {
            new { worksheet }.Must().NotBeNull();

            if (isHidden != null)
            {
                worksheet.IsRowColumnHeadersVisible = !(bool)isHidden;
            }
        }
    }
}
