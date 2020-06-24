// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NumericDataValidationTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using FakeItEasy;

    using FluentAssertions;

    using Xunit;

    public static partial class NumericDataValidationTest
    {
        [Fact]
        public static void Operand1Value___Should_be_same_as_Operand1NumericValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<NumericDataValidation>();

            // Act
            var actual = systemUnderTest.Operand1Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand1NumericValue);
        }

        [Fact]
        public static void Operand2Value___Should_be_same_as_Operand2NumericValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<NumericDataValidation>();

            // Act
            var actual = systemUnderTest.Operand2Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand2NumericValue);
        }
    }
}
