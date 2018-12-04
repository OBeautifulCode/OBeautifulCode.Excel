// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TypeConversionExtensionsTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;
    using System.Linq;

    using Aspose.Cells;

    using FluentAssertions;

    using OBeautifulCode.Enum.Recipes;

    using Xunit;

    public static class TypeConversionExtensionsTest
    {
        [Fact]
        public static void ToBorderType__Should_throw_ArgumentOutOfRangeException___When_parameter_borderEdges_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => BorderEdges.Unknown.ToBorderType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(BorderEdges.Unknown));
        }

        [Fact]
        public static void ToBorderType__Should_convert_borderEdges_to_a_BorderType___When_borderEdges_has_single_flag()
        {
            // Arrange
            var flags = EnumExtensions.GetIndividualFlags<BorderEdges>().Skip(1).ToList();
            var expected = flags.Select(_ => Enum.Parse(typeof(BorderType), _.ToString())).ToList();

            // Act
            var actual = flags.Select(_ => _.ToBorderType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToBorderType__Should_convert_borderEdges_to_a_BorderType___When_borderEdges_has_multiple_flags()
        {
            // Arrange
            var borderEdges1 = BorderEdges.DiagonalDown | BorderEdges.TopBorder;
            var borderEdges2 = BorderEdges.Outline;

            var expected1 = BorderType.DiagonalDown | BorderType.TopBorder;
            var expected2 = BorderType.BottomBorder | BorderType.TopBorder | BorderType.LeftBorder | BorderType.RightBorder;

            // Act
            var actual1 = borderEdges1.ToBorderType();
            var actual2 = borderEdges2.ToBorderType();

            // Act
            actual1.Should().Be(expected1);
            actual2.Should().Be(expected2);
        }

        [Fact]
        public static void ToCellBorderType__Should_convert_borderStyle_to_a_CellBorderType___When_called()
        {
            // Arrange
            var borderStyles = EnumExtensions.GetEnumValues<BorderStyle>();
            var expected = borderStyles.Select(_ => Enum.Parse(typeof(CellBorderType), _.ToString())).ToList();

            // Act
            var actual = borderStyles.Select(_ => _.ToCellBorderType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToValidationType__Should_convert_dataValidationKind_to_a_ValidationType___When_called()
        {
            // Arrange
            var dataValidationKinds = EnumExtensions.GetEnumValues<DataValidationKind>();
            var expected = dataValidationKinds.Select(_ => Enum.Parse(typeof(ValidationType), _.ToString())).ToList();

            // Act
            var actual = dataValidationKinds.Select(_ => _.ToValidationType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToOperatorType__Should_convert_dataValidationOperator_to_an_OperatorType___When_called()
        {
            // Arrange
            var dataValidationOperators = new[]
            {
                DataValidationOperator.Between,
                DataValidationOperator.EqualTo,
                DataValidationOperator.GreaterThan,
                DataValidationOperator.GreaterThanOrEqualTo,
                DataValidationOperator.LessThan,
                DataValidationOperator.LessThanOrEqualTo,
                DataValidationOperator.None,
                DataValidationOperator.NotBetween,
                DataValidationOperator.NotEqualTo,
            };

            var expected = new[]
            {
                OperatorType.Between,
                OperatorType.Equal,
                OperatorType.GreaterThan,
                OperatorType.GreaterOrEqual,
                OperatorType.LessThan,
                OperatorType.LessOrEqual,
                OperatorType.None,
                OperatorType.NotBetween,
                OperatorType.NotEqual,
            };

            // Act
            var actual = dataValidationOperators.Select(_ => _.ToOperatorType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToValidationAlertType__Should_convert_dataValidationErrorAlertStyle_to_a_ValidationAlertType___When_called()
        {
            // Arrange
            var dataValidationErrorAlertStyles = EnumExtensions.GetEnumValues<DataValidationErrorAlertStyle>();
            var expected = dataValidationErrorAlertStyles.Select(_ => Enum.Parse(typeof(ValidationAlertType), _.ToString())).ToList();

            // Act
            var actual = dataValidationErrorAlertStyles.Select(_ => _.ToValidationAlertType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }
    }
}
