// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SettingsHelper.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using Aspose.Cells;

    /// <summary>
    /// Convenience methods to manipulate various global Aspose settings.
    /// </summary>
    public static class SettingsHelper
    {
        /// <summary>
        /// Ensure that all operations that specify size in pixels, honor the number of pixels specified.
        /// </summary>
        /// <remarks>
        /// See <a href="https://forum.aspose.com/t/column-pixel-width-changes-based-on-monitor-resolution/190295/4" />.
        /// </remarks>
        public static void EnsureSizingOperationsHonorPixelsSpecified()
        {
            CellsHelper.DPI = 96;
        }
    }
}
