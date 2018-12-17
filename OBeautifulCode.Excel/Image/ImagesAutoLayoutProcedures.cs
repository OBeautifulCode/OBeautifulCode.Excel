// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ImagesAutoLayoutProcedures.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    /// <summary>
    /// Specifies which automatic layout procedures to apply to a set of images.
    /// </summary>
    [Flags]
    public enum ImagesAutoLayoutProcedures
    {
        /// <summary>
        /// Do not perform any automatic layout procedures.
        /// </summary>
        None = 0,

        /// <summary>
        /// Creates equal amounts of whitespace between the images.
        /// </summary>
        AutoSpace = 1,

        /// <summary>
        /// Aligns the images with each other.  If their orientation is <see cref="ImagesRelativeOrientation.Horizontal"/>
        /// then perform a vertical alignment.  If their orientation is <see cref="ImagesRelativeOrientation.Vertical"/>,
        /// then perform a horizontal alignment.
        /// </summary>
        AutoAlign = 2,

        /// <summary>
        /// Auto-space and auto-align.
        /// </summary>
        AutoSpaceAndAutoAlign = AutoSpace | AutoAlign,
    }
}
