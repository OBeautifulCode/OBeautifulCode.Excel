// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeExtensionsTest.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;
    using System.Linq;

    using Aspose.Cells;

    using FakeItEasy;

    using FluentAssertions;

    using Xunit;

    public static partial class RangeExtensionsTest
    {
        [Fact]
        public static void GetRowNumbers___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetRowNumbers(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetRowNumbers___Should_return_row_numbers_of_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetRowNumbers();
            var actual2 = range2.GetRowNumbers();
            var actual3 = range3.GetRowNumbers();
            var actual4 = range4.GetRowNumbers();

            // Assert
            actual1.Should().Equal(2);
            actual2.Should().Equal(2, 3, 4);
            actual3.Should().Equal(2);
            actual4.Should().Equal(2, 3, 4);
        }

        [Fact]
        public static void GetColumnNumbers___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetColumnNumbers(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetColumnNumbers___Should_return_column_numbers_of_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetColumnNumbers();
            var actual2 = range2.GetColumnNumbers();
            var actual3 = range3.GetColumnNumbers();
            var actual4 = range4.GetColumnNumbers();

            // Assert
            actual1.Should().Equal(2);
            actual2.Should().Equal(2);
            actual3.Should().Equal(2, 3, 4);
            actual4.Should().Equal(2, 3, 4);
        }

        [Fact]
        public static void GetCells___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetCells(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetCells___Should_return_cells_in_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetCells();
            var actual2 = range2.GetCells();
            var actual3 = range3.GetCells();
            var actual4 = range4.GetCells();

            // Assert
            actual1.Select(_ => _.Name).Should().Equal("B2");
            actual2.Select(_ => _.Name).Should().Equal("B2", "B3", "B4");
            actual3.Select(_ => _.Name).Should().Equal("B2", "C2", "D2");
            actual4.Select(_ => _.Name).Should().Equal("B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4", "D4");
        }

        [Fact]
        public static void GetCellRanges___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetCellRanges(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetCellRanges___Should_return_ranges_of_individual_cells_in_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetCellRanges();
            var actual2 = range2.GetCellRanges();
            var actual3 = range3.GetCellRanges();
            var actual4 = range4.GetCellRanges();

            // Assert
            actual1.Select(_ => _.RefersTo).Should().Equal("=Sheet1!$B$2");
            actual2.Select(_ => _.RefersTo).Should().Equal("=Sheet1!$B$2", "=Sheet1!$B$3", "=Sheet1!$B$4");
            actual3.Select(_ => _.RefersTo).Should().Equal("=Sheet1!$B$2", "=Sheet1!$C$2", "=Sheet1!$D$2");
            actual4.Select(_ => _.RefersTo).Should().Equal("=Sheet1!$B$2", "=Sheet1!$C$2", "=Sheet1!$D$2", "=Sheet1!$B$3", "=Sheet1!$C$3", "=Sheet1!$D$3", "=Sheet1!$B$4", "=Sheet1!$C$4", "=Sheet1!$D$4");
        }

        [Fact]
        public static void GetCellArea___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetCellArea(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetCellArea___Should_return_cell_area_of_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetCellArea();
            var actual2 = range2.GetCellArea();
            var actual3 = range3.GetCellArea();
            var actual4 = range4.GetCellArea();

            // Assert
            actual1.StartRow.Should().Be(1);
            actual1.StartColumn.Should().Be(1);
            actual1.EndRow.Should().Be(1);
            actual1.EndColumn.Should().Be(1);

            actual2.StartRow.Should().Be(1);
            actual2.StartColumn.Should().Be(1);
            actual2.EndRow.Should().Be(3);
            actual2.EndColumn.Should().Be(1);

            actual3.StartRow.Should().Be(1);
            actual3.StartColumn.Should().Be(1);
            actual3.EndRow.Should().Be(1);
            actual3.EndColumn.Should().Be(3);

            actual4.StartRow.Should().Be(1);
            actual4.StartColumn.Should().Be(1);
            actual4.EndRow.Should().Be(3);
            actual4.EndColumn.Should().Be(3);
        }

        [Fact]
        public static void GetUpperLeftmostCell___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => RangeExtensions.GetUpperLeftmostCell(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact]
        public static void GetUpperLeftmostCell___Should_return_upper_leftmost_cell___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var range1 = worksheet.Cells.CreateRange("B2");
            var range2 = worksheet.Cells.CreateRange("B2:B4");
            var range3 = worksheet.Cells.CreateRange("B2:D2");
            var range4 = worksheet.Cells.CreateRange("B2:D4");

            // Act
            var actual1 = range1.GetUpperLeftmostCell();
            var actual2 = range2.GetUpperLeftmostCell();
            var actual3 = range3.GetUpperLeftmostCell();
            var actual4 = range4.GetUpperLeftmostCell();

            // Assert
            actual1.Name.Should().Be("B2");
            actual2.Name.Should().Be("B2");
            actual3.Name.Should().Be("B2");
            actual4.Name.Should().Be("B2");
        }
    }
}
