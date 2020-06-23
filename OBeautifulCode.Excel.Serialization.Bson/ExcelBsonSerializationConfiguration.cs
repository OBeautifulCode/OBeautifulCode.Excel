// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelBsonSerializationConfiguration.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Bson
{
    using System.Collections.Generic;

    using OBeautifulCode.Serialization.Bson;

    using static System.FormattableString;

    /// <inheritdoc />
    public class ExcelBsonSerializationConfiguration : BsonSerializationConfigurationBase
    {
        /// <inheritdoc />
        protected override IReadOnlyCollection<TypeToRegisterForBson> TypesToRegisterForBson => new[]
        {
            typeof(Border).ToTypeToRegisterForBson(),
            typeof(Comment).ToTypeToRegisterForBson(),
            typeof(DataValidation).ToTypeToRegisterForBson(),
            typeof(DocumentProperties).ToTypeToRegisterForBson(),
            typeof(WorksheetProtection).ToTypeToRegisterForBson(),
            typeof(RangeStyle).ToTypeToRegisterForBson(),
            typeof(CellValueConditionalFormattingRule).ToTypeToRegisterForBson(),
            typeof(CellReference).ToTypeToRegisterForBson(),
            typeof(NamedCell).ToTypeToRegisterForBson(),
        };

        /// <inheritdoc />
        protected override IReadOnlyCollection<string> TypeToRegisterNamespacePrefixFilters => new[] { Invariant($"{nameof(OBeautifulCode)}.{nameof(Excel)}") };
    }
}
