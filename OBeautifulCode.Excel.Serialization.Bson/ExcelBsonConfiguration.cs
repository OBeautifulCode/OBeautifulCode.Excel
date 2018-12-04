// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelBsonConfiguration.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Bson
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;

    using MongoDB.Bson.Serialization;

    using Naos.Serialization.Bson;

    /// <inheritdoc />
    public class ExcelBsonConfiguration : BsonConfigurationBase
    {
        /// <inheritdoc />
        protected override IReadOnlyCollection<Type> TypesToAutoRegister => new[]
        {
            typeof(Border),
        };

        /// <inheritdoc />
        protected override void CustomConfiguration()
        {
            BsonSerializer.RegisterSerializer(typeof(Color), new ColorSerializer());
        }
    }
}
