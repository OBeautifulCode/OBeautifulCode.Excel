﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CustomFormatStringTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;

    using FakeItEasy;

    using FluentAssertions;

    using OBeautifulCode.AutoFakeItEasy;

    using Xunit;

    public static class CustomFormatStringTest
    {
        [Fact]
        public static void BuildCommonSeparatedThousandsWithDecimalsNumberFormat___Should_throw_ArgumentOutOfRangeException___When_parameter_numberOfDecimalPlaces_is_less_than_1()
        {
            // Arrange, Act
            var actual1 = Record.Exception(() => CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(0));
            var actual2 = Record.Exception(() => CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(A.Dummy<NegativeInteger>()));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("numberOfDecimalPlaces");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("numberOfDecimalPlaces");
        }

        [Fact]
        public static void BuildCommonSeparatedThousandsWithDecimalsNumberFormat___Should_throw_ArgumentOutOfRangeException___When_parameter_numberOfDecimalPlaces_is_greater_than_30()
        {
            // Arrange, Act
            var actual1 = Record.Exception(() => CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(31));
            var actual2 = Record.Exception(() => CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(A.Dummy<PositiveInteger>().ThatIs(_ => _ > 30)));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("numberOfDecimalPlaces");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("numberOfDecimalPlaces");
        }

        [Fact]
        public static void BuildCommonSeparatedThousandsWithDecimalsNumberFormat___Should_return_custom_format_string___When_called()
        {
            // Arrange
            var expected1 = "#,##0.0";
            var expected2 = "#,##0.00000";

            // Act
            var actual1 = CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(1);
            var actual2 = CustomFormatString.BuildCommonSeparatedThousandsWithDecimalsNumberFormat(5);

            // Assert
            actual1.Should().Be(expected1);
            actual2.Should().Be(expected2);
        }
    }
}
