// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorkbookProtection.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Type;

    /// <summary>
    /// The workbook protection configuration.
    /// </summary>
    public partial class WorkbookProtection : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the clear text password.
        /// </summary>
        public string ClearTextPassword { get; set; }
    }
}
