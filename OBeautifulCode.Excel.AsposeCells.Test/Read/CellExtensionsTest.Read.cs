// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensionsTest.Read.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;

    using Aspose.Cells;

    using FakeItEasy;

    using FluentAssertions;

    using Xunit;

    public static partial class CellExtensionsTest
    {
        [Fact]
        public static void GetRange___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.GetRange(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void GetRange___Should_return_range_corresponding_to_cell___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];
            var expected = worksheet.Cells.CreateRange("B2");

            // Act
            var actual = cell.GetRange();

            // Assert
            actual.RefersTo.Should().Be(expected.RefersTo);
        }

        [Fact]
        public static void GetWidthInPixels___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.GetWidthInPixels(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void GetWidthInPixels___Should_return_width_of_unmerged_cell___When_called()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];

            // Act
            var actual = cell.GetWidthInPixels();

            // Assert
            actual.Should().Be(Constants.DefaultColumnWidthInPixels);
        }

        [Fact]
        public static void GetWidthInPixels___Should_return_width_of_merged_cell___When_called()
        {
            // Arrange
            var worksheet1 = A.Dummy<Worksheet>();
            var range1 = worksheet1.Cells.CreateRange("B2:B4");
            range1.SetMergeCells(true);
            var cell1 = worksheet1.Cells[2, 1]; // B3

            var worksheet2 = A.Dummy<Worksheet>();
            var range2 = worksheet2.Cells.CreateRange("B2:D2");
            range2.SetMergeCells(true);
            var cell2A = worksheet2.Cells[1, 1]; // B2
            var cell2B = worksheet2.Cells[1, 3]; // D2

            var worksheet3 = A.Dummy<Worksheet>();
            var range3 = worksheet3.Cells.CreateRange("A1:C3");
            range3.SetMergeCells(true);
            var cell3 = worksheet3.Cells[1, 1]; // B2

            // Act
            var actual1 = cell1.GetWidthInPixels();
            var actual2A = cell2A.GetWidthInPixels();
            var actual2B = cell2B.GetWidthInPixels();
            var actual3 = cell3.GetWidthInPixels();

            // Assert
            cell1.IsMerged.Should().BeTrue();
            cell2A.IsMerged.Should().BeTrue();
            cell2B.IsMerged.Should().BeTrue();
            cell3.IsMerged.Should().BeTrue();

            actual1.Should().Be(Constants.DefaultColumnWidthInPixels);
            actual2A.Should().Be(Constants.DefaultColumnWidthInPixels * 3);
            actual2B.Should().Be(Constants.DefaultColumnWidthInPixels * 3);
            actual3.Should().Be(Constants.DefaultColumnWidthInPixels * 3);
        }
    }
}
