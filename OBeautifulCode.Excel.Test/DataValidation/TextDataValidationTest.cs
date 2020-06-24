// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TextDataValidationTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using FakeItEasy;

    using FluentAssertions;

    using Xunit;

    public static partial class TextDataValidationTest
    {
        [Fact]
        public static void Operand1Value___Should_be_same_as_Operand1TextValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<TextDataValidation>();

            // Act
            var actual = systemUnderTest.Operand1Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand1TextValue);
        }

        [Fact]
        public static void Operand2Value___Should_be_same_as_Operand2TextValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<TextDataValidation>();

            // Act
            var actual = systemUnderTest.Operand2Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand2TextValue);
        }
    }
}
