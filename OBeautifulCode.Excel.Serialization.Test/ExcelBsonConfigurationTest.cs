// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelBsonConfigurationTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Test
{
    using System.Drawing;

    using FakeItEasy;

    using FluentAssertions;

    using Naos.Serialization.Bson;

    using OBeautifulCode.Excel.Serialization.Bson;

    using Xunit;

    public static class ExcelBsonConfigurationTest
    {
        private static readonly NaosBsonSerializer Serializer = new NaosBsonSerializer(configurationType: typeof(ExcelBsonConfiguration));

        [Fact]
        public static void Configure___Should_not_throw___When_called_multiple_times()
        {
            // Arrange, Act
            var actual1 = Record.Exception(() => BsonConfigurationManager.Configure<ExcelBsonConfiguration>());
            var actual2 = Record.Exception(() => BsonConfigurationManager.Configure<ExcelBsonConfiguration>());

            // Assert
            actual1.Should().BeNull();
            actual2.Should().BeNull();
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_a_model_containing_a_Color_property___When_called()
        {
            // Arrange
            var expected1 = new ExcelTestModel { Color = Color.Empty };
            var expected2 = new ExcelTestModel { Color = A.Dummy<Color>() };
            var bytes1 = Serializer.SerializeToBytes(expected1);
            var bytes2 = Serializer.SerializeToBytes(expected2);

            // Act
            var actual1 = Serializer.Deserialize<ExcelTestModel>(bytes1);
            var actual2 = Serializer.Deserialize<ExcelTestModel>(bytes2);

            // Assert
            actual1.Color.Should().Be(expected1.Color);
            actual2.Color.Should().Be(expected2.Color);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_a_model_containing_a_nullable_Color_property___When_called()
        {
            // Arrange
            var expected1 = new ExcelTestModel();
            var expected2 = new ExcelTestModel { NullableColor = A.Dummy<Color>() };
            var bytes1 = Serializer.SerializeToBytes(expected1);
            var bytes2 = Serializer.SerializeToBytes(expected2);

            // Act
            var actual1 = Serializer.Deserialize<ExcelTestModel>(bytes1);
            var actual2 = Serializer.Deserialize<ExcelTestModel>(bytes2);

            // Assert
            actual1.NullableColor.Should().BeNull();
            actual2.NullableColor.Should().Be(expected2.NullableColor);
        }

        private class ExcelTestModel
        {
            public Color Color { get; set; }

            public Color? NullableColor { get; set; }
        }
    }
}
