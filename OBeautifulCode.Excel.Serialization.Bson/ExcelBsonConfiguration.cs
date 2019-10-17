// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelBsonConfiguration.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Bson
{
    using System;
    using System.Collections.Generic;

    using OBeautifulCode.Serialization.Bson;

    /// <inheritdoc />
    public class ExcelBsonConfiguration : BsonConfigurationBase
    {
        /// <inheritdoc />
        protected override IReadOnlyCollection<Type> TypesToAutoRegister => new[]
        {
            typeof(Border),
            typeof(Comment),
            typeof(DataValidation),
            typeof(DocumentProperties),
            typeof(WorksheetProtection),
            typeof(RangeStyle),
            typeof(CellValueConditionalFormattingRule),
            typeof(CellReference),
            typeof(NamedCell),
        };
    }
}
