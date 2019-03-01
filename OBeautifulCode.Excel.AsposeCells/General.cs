// --------------------------------------------------------------------------------------------------------------------
// <copyright file="General.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
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
        /// <returns>
        /// A new workbook.
        /// </returns>
        public static Workbook CreateStandardWorkbook()
        {
            EnsureSizingOperationsHonorPixelsSpecified();

            var result = new Workbook();

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
