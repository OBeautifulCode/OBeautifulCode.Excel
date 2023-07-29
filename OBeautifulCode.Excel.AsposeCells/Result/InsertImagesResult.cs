// --------------------------------------------------------------------------------------------------------------------
// <copyright file="InsertImagesResult.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Collections.Generic;
    using Aspose.Cells;
    using static System.FormattableString;
    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Result of inserting images.
    /// </summary>
    public class InsertImagesResult
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InsertImagesResult"/> class.
        /// </summary>
        /// <param name="containedWithinRange">The range of cells that the images are placed within or on top of.</param>
        /// <param name="pictureIndexes">The indices of the pictures inserted in <see cref="Worksheet.Pictures"/>, in the order specified when inserting.</param>
        public InsertImagesResult(
            Range containedWithinRange,
            IReadOnlyList<int> pictureIndexes)
        {
            if (containedWithinRange == null)
            {
                throw new ArgumentNullException(nameof(containedWithinRange));
            }

            if (pictureIndexes == null)
            {
                throw new ArgumentNullException(nameof(pictureIndexes));
            }

            if (pictureIndexes.Count == 0)
            {
                throw new ArgumentOutOfRangeException(Invariant($"{nameof(pictureIndexes)} is empty."));
            }

            this.ContainedWithinRange = containedWithinRange;
            this.PictureIndexes = pictureIndexes;
        }

        /// <summary>
        /// Gets the range of cells that the images are placed within or on top of.
        /// </summary>
        public Range ContainedWithinRange { get; }

        /// <summary>
        /// Gets the indices of the pictures inserted in <see cref="Worksheet.Pictures"/>, in the order specified when inserting.
        /// </summary>
        public IReadOnlyList<int> PictureIndexes { get; }
    }
}
