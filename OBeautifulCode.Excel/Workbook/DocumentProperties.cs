// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocumentProperties.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Collections.Generic;

    using OBeautifulCode.Type;

    /// <summary>
    /// Specifies the properties of a workbook document.
    /// </summary>
    public partial class DocumentProperties : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets a map of property kind to it's value.
        /// </summary>
        public IReadOnlyDictionary<BuiltInDocumentPropertyKind, string> BuiltInDocumentPropertyKindToValueMap { get; set; }
    }
}
