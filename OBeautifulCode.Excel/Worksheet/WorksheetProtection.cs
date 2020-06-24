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
    }
}
