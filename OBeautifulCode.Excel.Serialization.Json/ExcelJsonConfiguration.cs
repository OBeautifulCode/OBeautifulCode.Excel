// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelJsonConfiguration.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Json
{
    using System;
    using System.Collections.Generic;

    using Naos.Serialization.Json;

    /// <inheritdoc />
    public class ExcelJsonConfiguration : JsonConfigurationBase
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
