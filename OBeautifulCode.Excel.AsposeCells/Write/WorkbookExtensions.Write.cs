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
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static void RemoveDefaultWorksheet(
            this Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            workbook.Worksheets.RemoveAt(0);
        }

        /// <summary>
        /// Sets workbook document properties.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="documentProperties">The document properties to set.</param>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        public static void SetDocumentProperties(
            this Workbook workbook,
            DocumentProperties documentProperties)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (documentProperties != null)
            {
                var propertyKindToValueMap = documentProperties.BuiltInDocumentPropertyKindToValueMap;
                if (propertyKindToValueMap != null)
                {
                    foreach (var propertyKind in propertyKindToValueMap.Keys)
                    {
                        var propertyValue = propertyKindToValueMap[propertyKind];

                        if (propertyValue != null)
                        {
                            workbook.BuiltInDocumentProperties[propertyKind.ToBuiltInDocumentPropertyCollectionKey()].Value = propertyValue;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Applies the configured workbook protection.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="workbookProtection">The workbook protection.  If null then no protection is applied.</param>
        public static void SetWorkbookProtection(
            this Workbook workbook,
            WorkbookProtection workbookProtection)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (workbookProtection == null)
            {
                return;
            }

            workbook.Protect(ProtectionType.Structure, workbookProtection.ClearTextPassword);
        }
    }
}
