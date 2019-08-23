// --------------------------------------------------------------------------------------------------------------------
// <copyright file="General.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System.IO;

    using Aspose.Cells;

    /// <summary>
    /// Catch-all for higher level convenience methods such as configuring global settings
    /// and creating workbooks.
    /// </summary>
    public static class General
    {
        /// <summary>
        /// Creates a new workbook within a "standard" environment which ensures that
        /// the range of write-related helper methods in this library behave as expected.
        /// </summary>
        /// <param name="fileFormatType">The file format type to use.</param>
        /// <returns>
        /// A new workbook.
        /// </returns>
        public static Workbook CreateStandardWorkbook(
            FileFormatType fileFormatType = FileFormatType.Xlsx)
        {
            AsposeCellsLicense.ThrowIfNotRegistered();

            EnsureSizingOperationsHonorPixelsSpecified();

            var result = new Workbook(fileFormatType);

            return result;
        }

        /// <summary>
        /// Opens a workbook.
        /// </summary>
        /// <param name="stream">A stream with the contents of the workbook.</param>
        /// <param name="loadOptions">Optional load options to control how workbook is opened.</param>
        /// <returns>
        /// An open workbook.
        /// </returns>
        public static Workbook OpenWorkbook(
            Stream stream,
            LoadOptions loadOptions = null)
        {
            AsposeCellsLicense.ThrowIfNotRegistered();

            var workbookLoadOptions = loadOptions ?? new LoadOptions();
            var result = new Workbook(stream, workbookLoadOptions);

            return result;
        }

        /// <summary>
        /// Opens a workbook.
        /// </summary>
        /// <param name="filePath">The path to the workbook file.</param>
        /// <param name="loadOptions">Optional load options to control how workbook is opened.</param>
        /// <returns>
        /// An open workbook.
        /// </returns>
        public static Workbook OpenWorkbook(
            string filePath,
            LoadOptions loadOptions = null)
        {
            AsposeCellsLicense.ThrowIfNotRegistered();

            var workbookLoadOptions = loadOptions ?? new LoadOptions();
            var result = new Workbook(filePath, workbookLoadOptions);

            return result;
        }

        /// <summary>
        /// Ensure that all operations that specify size in pixels, honor the number of pixels specified.
        /// </summary>
        /// <remarks>
        /// See <a href="https://forum.aspose.com/t/column-pixel-width-changes-based-on-monitor-resolution/190295/4" />.
        /// We found that when we built worksheets on a machine where the resolution was set to 125%, the pixels
        /// we were specifying on sizing operations were not honored.  This happened when the worksheet creation
        /// was harnessed in an xUnit unit test.  In a console EXE this wasn't a problem.  This held true for
        /// Aspose.Cells 9.0 and 19.2.  At 100% resolution all combinations of harness and Aspose.Cells version worked fine.
        /// The issue is that, internally, Aspose.Cells uses inches to specify size and converts pixels to inches using
        /// the DPI setting.  The default is 96, but it seems when a harness has access to the Windows screen resolution
        /// it changes the DPI setting.
        /// </remarks>
        public static void EnsureSizingOperationsHonorPixelsSpecified()
        {
            CellsHelper.DPI = 96;
        }
    }
}
