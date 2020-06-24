// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SerializationConfigurationTypes.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using OBeautifulCode.Excel.Serialization.Bson;
    using OBeautifulCode.Excel.Serialization.Json;

    using OBeautifulCode.Serialization.Bson;
    using OBeautifulCode.Serialization.Json;

    public static class SerializationConfigurationTypes
    {
        public static BsonSerializationConfigurationType BsonSerializationConfigurationType => typeof(ExcelBsonSerializationConfiguration).ToBsonSerializationConfigurationType();

        public static JsonSerializationConfigurationType JsonSerializationConfigurationType => typeof(ExcelJsonSerializationConfiguration).ToJsonSerializationConfigurationType();
    }
}