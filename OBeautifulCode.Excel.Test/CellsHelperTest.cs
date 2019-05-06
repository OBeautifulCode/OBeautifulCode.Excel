// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellsHelperTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using FluentAssertions;

    using Xunit;

    public static class CellsHelperTest
    {
        [Fact]
        public static void GetColumnName___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_less_than_1()
        {
            // Arrange
            var columnNumbers = new[] { 0, -1, int.MinValue };

            // Act
            var actuals = columnNumbers.Select(_ => Record.Exception(() => CellsHelper.GetColumnName(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
            }
        }

        [Fact]
        public static void GetColumnName___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_greater_than_Constants_MaximumColumnNumber()
        {
            // Arrange
            var columnNumbers = new[] { Constants.MaximumColumnNumber + 1, int.MaxValue };

            // Act
            var actuals = columnNumbers.Select(_ => Record.Exception(() => CellsHelper.GetColumnName(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
            }
        }

        [Fact]
        public static void GetColumnName___Should_return_column_name___When_called()
        {
            // Arrange
            var columnNumberToExpectedColumnNameMap = new Dictionary<int, string>
            {
                { 1, "A" },
                { 2, "B" },
                { 26, "Z" },
                { 27, "AA" },
                { 52, "AZ" },
                { 53, "BA" },
                { 701, "ZY" },
                { 702, "ZZ" },
                { 703, "AAA" },
                { 704, "AAB" },
                { 10340, "OGR" },
                { 16384, "XFD" },
            };

            var expected = columnNumberToExpectedColumnNameMap.OrderBy(_ => _.Key).Select(_ => _.Value);

            // Act
            var actual = columnNumberToExpectedColumnNameMap.OrderBy(_ => _.Key).Select(_ => CellsHelper.GetColumnName(_.Key)).ToList();

            // Assert
            expected.Should().Equal(actual);
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentNullException___When_parameter_columnName_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellsHelper.GetColumnNumber(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("columnName");
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentException___When_parameter_columnName_is_white_space()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellsHelper.GetColumnNumber(" \r\n "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("columnName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentException___When_parameter_columnName_is_not_alphabetic()
        {
            var columnNames = new[] { "-", " A", "B ", "4" };

            // Act
            var actuals = columnNames.Select(_ => Record.Exception(() => CellsHelper.GetColumnNumber(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
                actual.Message.Should().Contain("columnName");
                actual.Message.Should().Contain("alphabetic");
            }
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentOutOfRangeException___When_parameter_columnName_contains_too_many_characters()
        {
            var columnNames = new[] { "abcd", "wierupweiqrupwqieurpwieorupiwqeurpwoierupioqewurpioeurpoiweurpoiweurpioweurioweuriewoureipwurpiweurepwirupweuirwepoiruwepoiru" };

            // Act
            var actuals = columnNames.Select(_ => Record.Exception(() => CellsHelper.GetColumnNumber(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNameLength");
                actual.Message.Should().Contain("3");
            }
        }

        [Fact]
        public static void GetColumnNumber___Should_return_column_numeric_corresponding_to_columnName___When_called()
        {
            var columnNameToExpectedColumnNumberMap = new Dictionary<string, int>
            {
                { "A", 1 },
                { "B", 2 },
                { "Z", 26 },
                { "AA", 27 },
                { "AZ", 52 },
                { "BA", 53 },
                { "ZY", 701 },
                { "ZZ", 702 },
                { "AAA", 703 },
                { "AAB", 704 },
                { "OGR", 10340 },
                { "XFD", 16384 },
            };

            var expected = columnNameToExpectedColumnNumberMap.OrderBy(_ => _.Key).Select(_ => _.Value);

            // Act
            var actual = columnNameToExpectedColumnNumberMap.OrderBy(_ => _.Key).Select(_ => CellsHelper.GetColumnNumber(_.Key)).ToList();

            // Assert
            expected.Should().Equal(actual);
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentOutOfRangeException___When_parameter_columnName_corresponds_to_column_number_that_is_greater_than_MaximumColumnNumber()
        {
            var columnNames = new[] { "XFE", "ZZZ" };

            // Act
            var actuals = columnNames.Select(_ => Record.Exception(() => CellsHelper.GetColumnNumber(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
            }
        }

        [Fact]
        public static void GetColumnNumber___Should_roundtrip_columnNumber_through_GetColumnName___When_called()
        {
            for (int expected = 1; expected <= Constants.MaximumColumnNumber; expected++)
            {
                // Arrange
                var columnName = CellsHelper.GetColumnName(expected);

                // Act
                var actual = CellsHelper.GetColumnNumber(columnName);

                // Assert
                actual.Should().Be(expected);
            }
        }
    }
}
