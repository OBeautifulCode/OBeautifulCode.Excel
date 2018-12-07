// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorkbookManipulationExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// Extensions methods to manipulate <see cref="Workbook"/>.
    /// </summary>
    public static class WorkbookManipulationExtensions
    {
        /// <summary>
        /// Adds a temporary worksheet to the workbook.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <returns>
        /// The temporary worksheet.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static Worksheet AddTemporaryWorksheet(
            this Workbook workbook)
        {
            new { workbook }.Must().NotBeNull();

            var worksheetName = Guid.NewGuid().ToString().Substring(0, 31);
            var worksheet = workbook.Worksheets.Add(worksheetName);
            return worksheet;
        }

        /// <summary>
        /// Removes the default worksheet.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static void RemoveDefaultWorksheet(
            this Workbook workbook)
        {
            new { workbook }.Must().NotBeNull();

            workbook.Worksheets.RemoveAt(0);
        }
    }
}
