// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorkbookExtensions.Write.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    /// <summary>
    /// Extensions methods on type <see cref="Workbook"/>.
    /// </summary>
    public static partial class WorkbookExtensions
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
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var worksheetName = Guid.NewGuid().ToString().Substring(0, 31);
            var worksheet = workbook.Worksheets.Add(worksheetName);
            return worksheet;
        }

        /// <summary>
        /// Removes the default worksheet.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <returns>
        /// The specified workbook with the default worksheet removed.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static Workbook RemoveDefaultWorksheet(
            this Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = workbook;

            result.Worksheets.RemoveAt(0);

            return result;
        }

        /// <summary>
        /// Sets workbook document properties.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <returns>
        /// The specified workbook with document properties set.
        /// </returns>
        /// <param name="documentProperties">The document properties to set.</param>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static Workbook SetDocumentProperties(
            this Workbook workbook,
            DocumentProperties documentProperties)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = workbook;

            if (documentProperties != null)
            {
                var builtInPropertyKindToValueMap = documentProperties.BuiltInDocumentPropertyKindToValueMap;
                if (builtInPropertyKindToValueMap != null)
                {
                    foreach (var propertyKind in builtInPropertyKindToValueMap.Keys)
                    {
                        var propertyValue = builtInPropertyKindToValueMap[propertyKind];

                        if (propertyValue != null)
                        {
                            result.BuiltInDocumentProperties[propertyKind.ToBuiltInDocumentPropertyCollectionKey()].Value = propertyValue;
                        }
                    }
                }

                var customPropertyNameToValueMap = documentProperties.CustomPropertyNameToValueMap;
                if (customPropertyNameToValueMap != null)
                {
                    foreach (var propertyName in customPropertyNameToValueMap.Keys)
                    {
                        var propertyValue = customPropertyNameToValueMap[propertyName];

                        if (propertyValue != null)
                        {
                            result.CustomDocumentProperties.Add(propertyName, propertyValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Applies the configured workbook protection.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="workbookProtection">The workbook protection.  If null then no protection is applied.</param>
        /// <returns>
        /// The specified workbook with the specified workbook protection applied.
        /// </returns>
        public static Workbook SetWorkbookProtection(
            this Workbook workbook,
            WorkbookProtection workbookProtection)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = workbook;

            if (workbookProtection == null)
            {
                return result;
            }

            result.Protect(ProtectionType.Structure, workbookProtection.ClearTextPassword);

            return result;
        }
    }
}
