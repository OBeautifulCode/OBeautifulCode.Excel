// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelJsonSerializationConfiguration.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Json
{
    using System.Collections.Generic;

    using OBeautifulCode.Serialization.Json;

    using static System.FormattableString;

    /// <inheritdoc />
    public class ExcelJsonSerializationConfiguration : JsonSerializationConfigurationBase
    {
        /// <inheritdoc />
        protected override IReadOnlyCollection<TypeToRegisterForJson> TypesToRegisterForJson => new[]
        {
            typeof(Border).ToTypeToRegisterForJson(),
            typeof(Comment).ToTypeToRegisterForJson(),
            typeof(DataValidation).ToTypeToRegisterForJson(),
            typeof(DocumentProperties).ToTypeToRegisterForJson(),
            typeof(WorksheetProtection).ToTypeToRegisterForJson(),
            typeof(WorkbookProtection).ToTypeToRegisterForJson(),
            typeof(RangeStyle).ToTypeToRegisterForJson(),
            typeof(CellValueConditionalFormattingRule).ToTypeToRegisterForJson(),
            typeof(CellReference).ToTypeToRegisterForJson(),
            typeof(NamedCell).ToTypeToRegisterForJson(),
        };

        /// <inheritdoc />
        protected override IReadOnlyCollection<string> TypeToRegisterNamespacePrefixFilters => new[] { Invariant($"{nameof(OBeautifulCode)}.{nameof(Excel)}") };
    }
}