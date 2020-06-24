// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetProtection.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Type;

    /// <summary>
    /// The worksheet protection configuration.
    /// </summary>
    public partial class WorksheetProtection : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the clear text password.
        /// </summary>
        public string ClearTextPassword { get; set; }

        /// <summary>
        /// Converts this workbook protection configuration to a worksheet protection configuration.
        /// </summary>
        /// <returns>
        /// The corresponding workbook protection.
        /// </returns>
        public WorkbookProtection ToWorkbookProtection()
        {
            var result = new WorkbookProtection
            {
                ClearTextPassword = this.ClearTextPassword,
            };

            return result;
        }
    }
}
