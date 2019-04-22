// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelJsonConfigurationTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Test
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;

    using FakeItEasy;

    using FluentAssertions;

    using Naos.Serialization.Domain;
    using Naos.Serialization.Json;

    using OBeautifulCode.Excel.Serialization.Json;

    using Xunit;

    public static class ExcelJsonConfigurationTest
    {
        private static readonly NaosJsonSerializer Serializer = new NaosJsonSerializer(typeof(ExcelTestJsonConfiguration));

        [Fact]
        public static void Configure___Should_not_throw___When_called_multiple_times()
        {
            // Arrange, Act
            var actual1 = Record.Exception(() => SerializationConfigurationManager.Configure<ExcelJsonConfiguration>());
            var actual2 = Record.Exception(() => SerializationConfigurationManager.Configure<ExcelJsonConfiguration>());

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

        private class ExcelTestJsonConfiguration : JsonConfigurationBase
        {
            public override IReadOnlyCollection<Type> DependentConfigurationTypes => new[] { typeof(ExcelJsonConfiguration) };

            protected override IReadOnlyCollection<Type> TypesToAutoRegister => new[] { typeof(ExcelTestModel) };
        }

        private class ExcelTestModel
        {
            public Color Color { get; set; }

            public Color? NullableColor { get; set; }
        }
    }
}
