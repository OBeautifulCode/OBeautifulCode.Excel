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
    }
}
