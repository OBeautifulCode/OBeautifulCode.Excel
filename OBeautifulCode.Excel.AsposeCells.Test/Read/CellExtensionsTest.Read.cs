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
        public static void GetWidthInPixels___Should_return_width_of_unmerged_cell___When_includeMergedCells_is_true()
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
        public static void GetWidthInPixels___Should_return_width_of_merged_cells___When_includeMergedCells_is_true()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var range1 = worksheet.Cells.CreateRange("B2:B4");
            range1.SetMergeCells(true);
            var cell1 = worksheet.Cells["B3"];

            var worksheet2 = A.Dummy<Worksheet>();
            var range2 = worksheet2.Cells.CreateRange("B2:D2");
            range2.SetMergeCells(true);
            var cell2A = worksheet2.Cells["B2"];
            var cell2B = worksheet2.Cells["D2"];

            var worksheet3 = A.Dummy<Worksheet>();
            var range3 = worksheet3.Cells.CreateRange("A1:C3");
            range3.SetMergeCells(true);
            var cell3 = worksheet3.Cells["B2"];

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

        [Fact]
        public static void GetWidthInPixels___Should_return_width_of_unmerged_cell___When_includeMergedCells_is_false()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];

            // Act
            var actual = cell.GetWidthInPixels(includeMergedCells: false);

            // Assert
            actual.Should().Be(Constants.DefaultColumnWidthInPixels);
        }

        [Fact]
        public static void GetWidthInPixels___Should_return_width_of_specified_cell___When_includeMergedCells_is_false()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var range1 = worksheet.Cells.CreateRange("B2:B4");
            range1.SetMergeCells(true);
            var cell1 = worksheet.Cells["B3"];

            var worksheet2 = A.Dummy<Worksheet>();
            var range2 = worksheet2.Cells.CreateRange("B2:D2");
            range2.SetMergeCells(true);
            var cell2A = worksheet2.Cells["B2"];
            var cell2B = worksheet2.Cells["D2"];

            var worksheet3 = A.Dummy<Worksheet>();
            var range3 = worksheet3.Cells.CreateRange("A1:C3");
            range3.SetMergeCells(true);
            var cell3 = worksheet3.Cells["B2"];

            // Act
            var actual1 = cell1.GetWidthInPixels(includeMergedCells: false);
            var actual2A = cell2A.GetWidthInPixels(includeMergedCells: false);
            var actual2B = cell2B.GetWidthInPixels(includeMergedCells: false);
            var actual3 = cell3.GetWidthInPixels(includeMergedCells: false);

            // Assert
            cell1.IsMerged.Should().BeTrue();
            cell2A.IsMerged.Should().BeTrue();
            cell2B.IsMerged.Should().BeTrue();
            cell3.IsMerged.Should().BeTrue();

            actual1.Should().Be(Constants.DefaultColumnWidthInPixels);
            actual2A.Should().Be(Constants.DefaultColumnWidthInPixels);
            actual2B.Should().Be(Constants.DefaultColumnWidthInPixels);
            actual3.Should().Be(Constants.DefaultColumnWidthInPixels);
        }

        [Fact]
        public static void GetHeightInPixels___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.GetHeightInPixels(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void GetHeightInPixels___Should_return_Height_of_unmerged_cell___When_includeMergedCells_is_true()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];

            // Act
            var actual = cell.GetHeightInPixels();

            // Assert
            actual.Should().Be(Constants.DefaultRowHeightInPixels);
        }

        [Fact]
        public static void GetHeightInPixels___Should_return_Height_of_merged_cells___When_includeMergedCells_is_true()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var range1 = worksheet.Cells.CreateRange("B2:B4");
            range1.SetMergeCells(true);
            var cell1A = worksheet.Cells["B2"];
            var cell1B = worksheet.Cells["B4"];

            var worksheet2 = A.Dummy<Worksheet>();
            var range2 = worksheet2.Cells.CreateRange("B2:D2");
            range2.SetMergeCells(true);
            var cell2 = worksheet2.Cells["B2"];

            var worksheet3 = A.Dummy<Worksheet>();
            var range3 = worksheet3.Cells.CreateRange("A1:C3");
            range3.SetMergeCells(true);
            var cell3 = worksheet3.Cells["B2"];

            // Act
            var actual1A = cell1A.GetHeightInPixels();
            var actual1B = cell1B.GetHeightInPixels();
            var actual2 = cell2.GetHeightInPixels();
            var actual3 = cell3.GetHeightInPixels();

            // Assert
            cell1A.IsMerged.Should().BeTrue();
            cell1B.IsMerged.Should().BeTrue();
            cell2.IsMerged.Should().BeTrue();
            cell3.IsMerged.Should().BeTrue();

            actual1A.Should().Be(Constants.DefaultRowHeightInPixels * 3);
            actual1B.Should().Be(Constants.DefaultRowHeightInPixels * 3);
            actual2.Should().Be(Constants.DefaultRowHeightInPixels);
            actual3.Should().Be(Constants.DefaultRowHeightInPixels * 3);
        }

        [Fact]
        public static void GetHeightInPixels___Should_return_Height_of_unmerged_cell___When_includeMergedCells_is_false()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];

            // Act
            var actual = cell.GetHeightInPixels(includeMergedCells: false);

            // Assert
            actual.Should().Be(Constants.DefaultRowHeightInPixels);
        }

        [Fact]
        public static void GetHeightInPixels___Should_return_Height_of_specified_cell___When_includeMergedCells_is_false()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var range1 = worksheet.Cells.CreateRange("B2:B4");
            range1.SetMergeCells(true);
            var cell1A = worksheet.Cells["B2"];
            var cell1B = worksheet.Cells["B4"];

            var worksheet2 = A.Dummy<Worksheet>();
            var range2 = worksheet2.Cells.CreateRange("B2:D2");
            range2.SetMergeCells(true);
            var cell2 = worksheet2.Cells["B2"];

            var worksheet3 = A.Dummy<Worksheet>();
            var range3 = worksheet3.Cells.CreateRange("A1:C3");
            range3.SetMergeCells(true);
            var cell3 = worksheet3.Cells["B2"];

            // Act
            var actual1A = cell1A.GetHeightInPixels(includeMergedCells: false);
            var actual1B = cell1B.GetHeightInPixels(includeMergedCells: false);
            var actual2 = cell2.GetHeightInPixels(includeMergedCells: false);
            var actual3 = cell3.GetHeightInPixels(includeMergedCells: false);

            // Assert
            cell1A.IsMerged.Should().BeTrue();
            cell1B.IsMerged.Should().BeTrue();
            cell2.IsMerged.Should().BeTrue();
            cell3.IsMerged.Should().BeTrue();

            actual1A.Should().Be(Constants.DefaultRowHeightInPixels);
            actual1B.Should().Be(Constants.DefaultRowHeightInPixels);
            actual2.Should().Be(Constants.DefaultRowHeightInPixels);
            actual3.Should().Be(Constants.DefaultRowHeightInPixels);
        }

        [Fact]
        public static void GetRowNumber___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.GetRowNumber(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void GetRowNumber___Should_return_row_number_of_cell___When_called()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells["C9"];

            // Act
            var actual = cell.GetRowNumber();

            // Assert
            actual.Should().Be(9);
        }

        [Fact]
        public static void GetColumnNumber___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.GetColumnNumber(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void GetColumnNumber___Should_return_column_number_of_cell___When_called()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells["C9"];

            // Act
            var actual = cell.GetColumnNumber();

            // Assert
            actual.Should().Be(3);
        }

        [Fact]
        public static void ToCellReference___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellExtensions.ToCellReference(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void ToCellReference___Should_get_the_CellReference_corresponding_to_the_specified_cell___When_called()
        {
            // Arrange,
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells["C9"];
            var expected = new CellReference("Sheet1", 9, 3);

            // Act
            var actual = cell.ToCellReference();

            // Assert
            actual.Should().Be(expected);
        }
    }
}
