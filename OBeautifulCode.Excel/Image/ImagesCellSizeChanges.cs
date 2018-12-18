// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ImagesCellSizeChanges.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    /// <summary>
    /// Specifies the changes to make to the size of a cell to fit a set of images.
    /// </summary>
    [Flags]
    public enum ImagesCellSizeChanges
    {
        /// <summary>
        /// Do not change the size of the cell.  Images may overflow the cell.
        /// </summary>
        None = 0,

        /// <summary>
        /// Expands the size of the row to fit the images.
        /// </summary>
        ExpandRowToFitImages = 1,

        /// <summary>
        /// Expands the size of the column to fit the images.
        /// </summary>
        ExpandColumnToFitImages = 2,

        /// <summary>
        /// Resize all rows that overlap with the image to a fixed height.
        /// </summary>
        ResizeRowsToFixedHeight = 4,

        /// <summary>
        /// Resize all column that overlap with the image to a fixed height.
        /// </summary>
        ResizeColumnsToFixedWidth = 8,

        /// <summary>
        /// Expands the size of the row and the column to fit the images.
        /// </summary>
        ExpandRowAndColumnToFitImages = ExpandRowToFitImages | ExpandColumnToFitImages,
    }
}
