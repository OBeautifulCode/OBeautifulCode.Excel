// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidationErrorAlertStyle.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    /// <summary>
    /// Determines the style of error alert to show on a data validation.
    /// </summary>
    public enum DataValidationErrorAlertStyle
    {
        /// <summary>
        /// Unknown (default).
        /// </summary>
        Unknown,

        /// <summary>
        /// Show an informational alert.
        /// </summary>
        Information,

        /// <summary>
        /// Show a stopping alert.  User won't be able to continue until the data validates.
        /// </summary>
        Stop,

        /// <summary>
        /// Show a warning alert.  User will be given the option to continue with the invalid data or not.
        /// </summary>
        Warning,
    }
}
