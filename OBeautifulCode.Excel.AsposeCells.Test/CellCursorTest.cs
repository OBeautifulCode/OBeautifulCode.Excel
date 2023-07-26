// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellCursorTest.cs" company="OBeautifulCode">
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

    using OBeautifulCode.AutoFakeItEasy;

    using Xunit;

    public static class CellCursorTest
    {
        private const string DefaultMarkerNamePrefix = "marker-";

        private const int StartRowNumber = 3;
        private const int StartColumnNumber = 4;

        private const int MoveDownBy = 4;
        private const int MoveRightBy = 8;
        private const int MoveLeftBy = 6;
        private const int MoveUpBy = 2;

        private const int MaxRowNumber = StartRowNumber + MoveDownBy;
        private const int MaxColumnNumber = StartColumnNumber + MoveRightBy;

        private const int CurrentRowNumber = StartRowNumber + MoveDownBy - MoveUpBy;
        private const int CurrentColumnNumber = StartColumnNumber + MoveRightBy - MoveLeftBy;

        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_worksheet_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new CellCursor(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheet");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_rowNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual1 = Record.Exception(() => new CellCursor(worksheet, rowNumber: 0));
            var actual2 = Record.Exception(() => new CellCursor(worksheet, rowNumber: A.Dummy<NegativeInteger>()));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("rowNumber");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual1 = Record.Exception(() => new CellCursor(worksheet, columnNumber: 0));
            var actual2 = Record.Exception(() => new CellCursor(worksheet, columnNumber: A.Dummy<NegativeInteger>()));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("columnNumber");
        }

        [Fact]
        public static void Worksheet___Should_return_same_worksheet_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var systemUnderTest = new CellCursor(worksheet);

            // Act
            var actual = systemUnderTest.Worksheet;

            // Assert
            actual.Should().BeSameAs(worksheet);
        }

        [Fact]
        public static void RowNumber___Should_return_same_rowNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var rowNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, rowNumber: rowNumber);

            // Act
            var actual = systemUnderTest.RowNumber;

            // Assert
            actual.Should().Be(rowNumber);
        }

        [Fact]
        public static void ColumnNumber___Should_return_same_columnNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var columnNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, columnNumber: columnNumber);

            // Act
            var actual = systemUnderTest.ColumnNumber;

            // Assert
            actual.Should().Be(columnNumber);
        }

        [Fact]
        public static void StartRowNumber___Should_return_same_rowNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var rowNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, rowNumber: rowNumber);

            // Act
            var actual = systemUnderTest.StartRowNumber;

            // Assert
            actual.Should().Be(rowNumber);
        }

        [Fact]
        public static void StartColumnNumber___Should_return_same_columnNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var columnNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, columnNumber: columnNumber);

            // Act
            var actual = systemUnderTest.StartColumnNumber;

            // Assert
            actual.Should().Be(columnNumber);
        }

        [Fact]
        public static void MaxRowNumber___Should_return_same_rowNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var rowNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, rowNumber: rowNumber);

            // Act
            var actual = systemUnderTest.MaxRowNumber;

            // Assert
            actual.Should().Be(rowNumber);
        }

        [Fact]
        public static void MaxColumnNumber___Should_return_same_columnNumber_passed_to_constructor___When_getting()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var columnNumber = A.Dummy<PositiveInteger>();
            var systemUnderTest = new CellCursor(worksheet, columnNumber: columnNumber);

            // Act
            var actual = systemUnderTest.MaxColumnNumber;

            // Assert
            actual.Should().Be(columnNumber);
        }

        [Fact]
        public static void Cell___Should_return_cell_under_cursor___When_getting()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.Cell;

            // Assert
            actual.ToCellReference().A1Reference.Should().Be("F5");
        }

        [Fact]
        public static void CellRange___Should_return_range_of_cell_under_cursor___When_getting()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.CellRange;

            // Assert
            actual.RefersTo.Should().Be("=Sheet1!$F$5");
        }

        [Fact]
        public static void CanvassedRowRange___Should_return_canvassed_range_of_row_under_cursor___When_getting()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.CanvassedRowRange;

            // Assert
            actual.RefersTo.Should().Be("=Sheet1!$D$5:$L$5");
        }

        [Fact]
        public static void CanvassedColumnRange___Should_return_canvassed_range_of_column_under_cursor___When_getting()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.CanvassedColumnRange;

            // Assert
            actual.RefersTo.Should().Be("=Sheet1!$F$3:$F$7");
        }

        [Fact]
        public static void CanvassedRange___Should_return_canvassed_range___When_getting()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.CanvassedRange;

            // Assert
            actual.RefersTo.Should().Be("=Sheet1!$D$3:$L$7");
        }

        [Fact]
        public static void CloneWithWorksheet___Should_throw_ArgumentNullException___When_parameter_worksheet_is_null()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = Record.Exception(() => systemUnderTest.CloneWithWorksheet(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheet");
        }

        [Fact]
        public static void CloneWithWorksheet___Should_clone_the_CellCursor_and_replace_worksheet_in_all_properties___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var worksheet = A.Dummy<Worksheet>();

            var worksheetName = worksheet.Name;
            var expectedA1 = new CellReference(worksheetName, 1, 1);
            var expectedB1 = new CellReference(worksheetName, 1, 2);
            var expectedC1 = new CellReference(worksheetName, 1, 3);
            var expectedA2 = new CellReference(worksheetName, 2, 1);
            var expectedB2 = new CellReference(worksheetName, 2, 2);
            var expectedC2 = new CellReference(worksheetName, 2, 3);
            var expectedA3 = new CellReference(worksheetName, 3, 1);
            var expectedB3 = new CellReference(worksheetName, 3, 2);
            var expectedC3 = new CellReference(worksheetName, 3, 3);

            // Act
            var actual = systemUnderTest.CloneWithWorksheet(worksheet);

            // Assert
            actual.RowNumber.Should().Be(systemUnderTest.RowNumber);
            actual.ColumnNumber.Should().Be(systemUnderTest.ColumnNumber);
            actual.StartRowNumber.Should().Be(systemUnderTest.StartRowNumber);
            actual.StartColumnNumber.Should().Be(systemUnderTest.StartColumnNumber);
            actual.MaxRowNumber.Should().Be(systemUnderTest.MaxRowNumber);
            actual.MaxColumnNumber.Should().Be(systemUnderTest.MaxColumnNumber);
            actual.Worksheet.Should().Be(worksheet);

            actual.GetMarkedCellReference("marker-A1").Should().Be(expectedA1);
            actual.GetMarkedCellReference("marker-B1").Should().Be(expectedB1);
            actual.GetMarkedCellReference("marker-C1").Should().Be(expectedC1);
            actual.GetMarkedCellReference("marker-A2").Should().Be(expectedA2);
            actual.GetMarkedCellReference("marker-B2").Should().Be(expectedB2);
            actual.GetMarkedCellReference("marker-C2").Should().Be(expectedC2);
            actual.GetMarkedCellReference("marker-A3").Should().Be(expectedA3);
            actual.GetMarkedCellReference("marker-B3").Should().Be(expectedB3);
            actual.GetMarkedCellReference("marker-C3").Should().Be(expectedC3);
        }

        [Fact]
        public static void Reset___Should_reset_the_cursor_to_the_starting_row_and_starting_column___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.Reset();

            // Assert
            systemUnderTest.RowNumber.Should().Be(StartRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void Reset___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.Reset();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void ResetRow___Should_reset_the_cursor_to_the_starting_row___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.ResetRow();

            // Assert
            systemUnderTest.RowNumber.Should().Be(StartRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void ResetRow___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.ResetRow();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void ResetColumn___Should_reset_the_cursor_to_the_starting_column___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.ResetColumn();

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void ResetColumn___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.ResetColumn();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveToBottomRightOfCanvas___Should_move_cursor_to_MaxRow_and_MaxColumn___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveToBottomRightOfCanvas();

            // Assert
            systemUnderTest.RowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveToBottomRightOfCanvas___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveToBottomRightOfCanvas();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveDownToMaxRow___Should_move_cursor_to_MaxRow___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveDownToMaxRow();

            // Assert
            systemUnderTest.RowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveDownToMaxRow___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveDownToMaxRow();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveRightToMaxColumn___Should_move_cursor_to_MaxColumns___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveRightToMaxColumn();

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveRightToMaxColumn___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveRightToMaxColumn();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveDown___Should_throw_ArgumentOutOfRangeException___When_parameter_by_is_less_than_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveDown(A.Dummy<NegativeInteger>()));

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain("by");
        }

        [Fact]
        public static void MoveDown___Should_not_move_cursor___When_parameter_by_is_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveDown(0);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveDown___Should_move_cursor_down_without_changing_size_of_canvas___When_moving_down_but_not_exceeding_MaxRow()
        {
            // Arrange
            var systemUnderTest1 = GetCursorWithUniqueRowAndColumnNumbers();
            var systemUnderTest2 = GetCursorWithUniqueRowAndColumnNumbers();

            var by1 = MaxRowNumber - CurrentRowNumber - 1;
            var by2 = MaxRowNumber - CurrentRowNumber;

            // Act
            systemUnderTest1.MoveDown(by1);
            systemUnderTest2.MoveDown(by2);

            // Assert
            systemUnderTest1.RowNumber.Should().Be(CurrentRowNumber + by1);
            systemUnderTest1.ColumnNumber.Should().Be(CurrentColumnNumber);
            systemUnderTest1.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest1.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest1.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest1.MaxColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest2.RowNumber.Should().Be(CurrentRowNumber + by2);
            systemUnderTest2.ColumnNumber.Should().Be(CurrentColumnNumber);
            systemUnderTest2.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest2.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest2.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest2.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveDown___Should_move_cursor_down_and_change_size_of_canvas___When_moving_past_MaxRow()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();
            var by = MaxRowNumber - CurrentRowNumber + 1;

            // Act
            systemUnderTest.MoveDown(by);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber + by);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);
            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber + 1);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveDown___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveDown();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveUp___Should_throw_ArgumentOutOfRangeException___When_parameter_by_is_less_than_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveUp(A.Dummy<NegativeInteger>()));

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain("by");
        }

        [Fact]
        public static void MoveUp___Should_not_move_cursor___When_parameter_by_is_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveUp(0);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveUp___Should_throw_InvalidOperationException___When_moving_before_StartRow()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();
            var by = CurrentRowNumber - StartRowNumber + 1;

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveUp(by));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("because it would place cursor above start row");
        }

        [Fact]
        public static void MoveUp___Should_move_cursor_up_without_changing_size_of_canvas___When_moving_up_but_staying_at_or_below_StartRow()
        {
            // Arrange
            var systemUnderTest1 = GetCursorWithUniqueRowAndColumnNumbers();
            var systemUnderTest2 = GetCursorWithUniqueRowAndColumnNumbers();

            var by1 = 1;
            var by2 = CurrentRowNumber - StartRowNumber;

            // Act
            systemUnderTest1.MoveUp(by1);
            systemUnderTest2.MoveUp(by2);

            // Assert
            systemUnderTest1.RowNumber.Should().Be(CurrentRowNumber - 1);
            systemUnderTest1.ColumnNumber.Should().Be(CurrentColumnNumber);
            systemUnderTest1.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest1.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest1.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest1.MaxColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest2.RowNumber.Should().Be(StartRowNumber);
            systemUnderTest2.ColumnNumber.Should().Be(CurrentColumnNumber);
            systemUnderTest2.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest2.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest2.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest2.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveUp___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveUp();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveRight___Should_throw_ArgumentOutOfRangeException___When_parameter_by_is_less_than_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveRight(A.Dummy<NegativeInteger>()));

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain("by");
        }

        [Fact]
        public static void MoveRight___Should_not_move_cursor___When_parameter_by_is_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveRight(0);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveRight___Should_move_cursor_right_without_changing_size_of_canvas___When_moving_right_but_not_exceeding_MaxColumn()
        {
            // Arrange
            var systemUnderTest1 = GetCursorWithUniqueRowAndColumnNumbers();
            var systemUnderTest2 = GetCursorWithUniqueRowAndColumnNumbers();

            var by1 = MaxColumnNumber - CurrentColumnNumber - 1;
            var by2 = MaxColumnNumber - CurrentColumnNumber;

            // Act
            systemUnderTest1.MoveRight(by1);
            systemUnderTest2.MoveRight(by2);

            // Assert
            systemUnderTest1.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest1.ColumnNumber.Should().Be(CurrentColumnNumber + by1);
            systemUnderTest1.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest1.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest1.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest1.MaxColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest2.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest2.ColumnNumber.Should().Be(CurrentColumnNumber + by2);
            systemUnderTest2.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest2.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest2.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest2.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveRight___Should_move_cursor_right_and_change_size_of_canvas___When_moving_past_MaxColumn()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();
            var by = MaxColumnNumber - CurrentColumnNumber + 1;

            // Act
            systemUnderTest.MoveRight(by);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber + by);
            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber + 1);
        }

        [Fact]
        public static void MoveRight___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveRight();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void MoveLeft___Should_throw_ArgumentOutOfRangeException___When_parameter_by_is_less_than_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveLeft(A.Dummy<NegativeInteger>()));

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain("by");
        }

        [Fact]
        public static void MoveLeft___Should_not_move_cursor___When_parameter_by_is_0()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            systemUnderTest.MoveLeft(0);

            // Assert
            systemUnderTest.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest.ColumnNumber.Should().Be(CurrentColumnNumber);

            systemUnderTest.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest.StartColumnNumber.Should().Be(StartColumnNumber);

            systemUnderTest.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveLeft___Should_throw_InvalidOperationException___When_moving_before_StartColumn()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();
            var by = CurrentColumnNumber - StartColumnNumber + 1;

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveLeft(by));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("because it would place cursor to the left of start column");
        }

        [Fact]
        public static void MoveLeft___Should_move_cursor_left_without_changing_size_of_canvas___When_moving_left_but_staying_at_or_past_StartColumn()
        {
            // Arrange
            var systemUnderTest1 = GetCursorWithUniqueRowAndColumnNumbers();
            var systemUnderTest2 = GetCursorWithUniqueRowAndColumnNumbers();

            var by1 = 1;
            var by2 = CurrentColumnNumber - StartColumnNumber;

            // Act
            systemUnderTest1.MoveLeft(by1);
            systemUnderTest2.MoveLeft(by2);

            // Assert
            systemUnderTest1.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest1.ColumnNumber.Should().Be(CurrentColumnNumber - 1);
            systemUnderTest1.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest1.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest1.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest1.MaxColumnNumber.Should().Be(MaxColumnNumber);

            systemUnderTest2.RowNumber.Should().Be(CurrentRowNumber);
            systemUnderTest2.ColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest2.StartRowNumber.Should().Be(StartRowNumber);
            systemUnderTest2.StartColumnNumber.Should().Be(StartColumnNumber);
            systemUnderTest2.MaxRowNumber.Should().Be(MaxRowNumber);
            systemUnderTest2.MaxColumnNumber.Should().Be(MaxColumnNumber);
        }

        [Fact]
        public static void MoveLeft___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWithUniqueRowAndColumnNumbers();

            // Act
            var actual = systemUnderTest.MoveLeft();

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void AddMarker___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.AddMarker(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void AddMarker___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.AddMarker("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void AddMarker___Should_add_the_specified_marker___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            systemUnderTest.Reset();
            systemUnderTest.MoveRight().MoveDown();

            var markerName = A.Dummy<string>();

            // Act
            systemUnderTest.AddMarker(markerName);

            // Assert
            systemUnderTest.HasMarker(markerName).Should().BeTrue();
            systemUnderTest.GetMarkedCell(markerName).ToCellReference().A1Reference.Should().Be("B2");
            systemUnderTest.GetMarkedCells(markerName).SingleOrDefault().ToCellReference().A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void AddMarker___Should_not_add_the_same_marker_twice_to_the_same_cell___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            systemUnderTest.Reset();
            systemUnderTest.MoveRight().MoveDown();

            var markerName = A.Dummy<string>();
            systemUnderTest.AddMarker(markerName);

            // Act
            systemUnderTest.AddMarker(markerName);

            // Assert
            systemUnderTest.HasMarker(markerName).Should().BeTrue();
            systemUnderTest.GetMarkedCell(markerName).ToCellReference().A1Reference.Should().Be("B2");
            systemUnderTest.GetMarkedCells(markerName).SingleOrDefault().ToCellReference().A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_ArgumentNullException___When_parameter_sourceMarkerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers(null, "target"));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("sourceMarkerName");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_ArgumentException___When_parameter_sourceMarkerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers("  \r\n  ", "target"));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("sourceMarkerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_ArgumentNullException___When_parameter_targetMarkerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers("source", null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("targetMarkerName");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_ArgumentException___When_parameter_targetMarkerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers("source", "  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("targetMarkerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_InvalidOperationException___When_sourceMarker_does_not_exist()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();
            systemUnderTest.AddMarker("target");

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers(A.Dummy<string>(), "target"));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("source marker does not exist");
        }

        [Fact]
        public static void MergeMarkers___Should_throw_InvalidOperationException___When_targetMarker_does_not_exist()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();
            systemUnderTest.AddMarker("source");

            // Act
            var actual = Record.Exception(() => systemUnderTest.MergeMarkers("source", A.Dummy<string>()));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("target marker does not exist");
        }

        [Fact]
        public static void MergeMarkers___Should_merge_source_into_target___When_there_is_no_overlap_in_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var sourceMarkerName = "source";
            var targetMarkerName = "target";

            systemUnderTest.Reset();
            systemUnderTest.AddMarker(sourceMarkerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(sourceMarkerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(targetMarkerName);

            // Act
            var actual = systemUnderTest.MergeMarkers(sourceMarkerName, targetMarkerName);

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
            actual.HasMarker(sourceMarkerName).Should().BeFalse();
            actual.GetMarkedCells(targetMarkerName).Select(_ => _.ToCellReference().A1Reference).Should().BeEquivalentTo("A1", "B2", "C3");
        }

        [Fact]
        public static void MergeMarkers___Should_merge_source_into_target___When_there_is_overlap_in_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var sourceMarkerName = "source";
            var targetMarkerName = "target";

            systemUnderTest.Reset();
            systemUnderTest.AddMarker(sourceMarkerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(sourceMarkerName);
            systemUnderTest.AddMarker(targetMarkerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(targetMarkerName);

            // Act
            var actual = systemUnderTest.MergeMarkers(sourceMarkerName, targetMarkerName);

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
            actual.HasMarker(sourceMarkerName).Should().BeFalse();
            actual.GetMarkedCells(targetMarkerName).Select(_ => _.ToCellReference().A1Reference).Should().BeEquivalentTo("A1", "B2", "C3");
        }

        [Fact]
        public static void RemoveAllMarkers___Should_remove_all_markers___When_called()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            systemUnderTest1.RemoveAllMarkers();
            systemUnderTest2.RemoveAllMarkers();

            // Assert
            // nothing that we can really test for systemUnderTest1
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}A1").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}A2").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}A3").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}B1").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}B2").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}B3").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}C1").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}C2").Should().BeFalse();
            systemUnderTest2.HasMarker($"{DefaultMarkerNamePrefix}C3").Should().BeFalse();
        }

        [Fact]
        public static void RemoveMarker___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.RemoveMarker(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void RemoveMarker___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.RemoveMarker("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void RemoveMarker___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.RemoveMarker($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        [Fact]
        public static void RemoveMarker___Should_remove_the_specified_marker___When_only_one_cell_is_marked_with_that_marker()
        {
            // Arrange
            var markerNamePrefix = A.Dummy<string>();
            var systemUnderTest = GetCursorWith3x3MarkedCanvas(markerNamePrefix);

            var markerName = $"{markerNamePrefix}B2";
            systemUnderTest.HasMarker(markerName).Should().BeTrue();

            // Act
            systemUnderTest.RemoveMarker(markerName);

            // Assert
            systemUnderTest.HasMarker(markerName).Should().BeFalse();
        }

        [Fact]
        public static void RemoveMarker___Should_remove_the_specified_marker___When_multiple_cells_are_marked_with_that_marker()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();

            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);

            // Act
            systemUnderTest.RemoveMarker(markerName);

            // Assert
            systemUnderTest.HasMarker(markerName).Should().BeFalse();
        }

        [Fact]
        public static void HasMarker___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.HasMarker(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void HasMarker___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.HasMarker("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void HasMarker___Should_return_false___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = systemUnderTest1.HasMarker(A.Dummy<string>());
            var actual2 = systemUnderTest2.HasMarker(A.Dummy<string>());

            // Assert
            actual1.Should().BeFalse();
            actual2.Should().BeFalse();
        }

        [Fact]
        public static void HasMarker___Should_return_true___When_marker_exists_on_single_cell()
        {
            // Arrange
            var markerPrefix = A.Dummy<string>();
            var systemUnderTest = GetCursorWith3x3MarkedCanvas(markerPrefix);

            // Act
            var actual1 = systemUnderTest.HasMarker($"{markerPrefix}A1");
            var actual2 = systemUnderTest.HasMarker($"{markerPrefix}A2");
            var actual3 = systemUnderTest.HasMarker($"{markerPrefix}A3");
            var actual4 = systemUnderTest.HasMarker($"{markerPrefix}B1");
            var actual5 = systemUnderTest.HasMarker($"{markerPrefix}B2");
            var actual6 = systemUnderTest.HasMarker($"{markerPrefix}B3");
            var actual7 = systemUnderTest.HasMarker($"{markerPrefix}C1");
            var actual8 = systemUnderTest.HasMarker($"{markerPrefix}C2");
            var actual9 = systemUnderTest.HasMarker($"{markerPrefix}C3");

            // Assert
            actual1.Should().BeTrue();
            actual2.Should().BeTrue();
            actual3.Should().BeTrue();
            actual4.Should().BeTrue();
            actual5.Should().BeTrue();
            actual6.Should().BeTrue();
            actual7.Should().BeTrue();
            actual8.Should().BeTrue();
            actual9.Should().BeTrue();
        }

        [Fact]
        public static void HasMarker___Should_return_true___When_marker_exists_on_multiple_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();

            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);

            // Act
            var actual = systemUnderTest.HasMarker(markerName);

            // Assert
            actual.Should().BeTrue();
        }

        [Fact]
        public static void GetMarkedRange___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedRange(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void GetMarkedRange___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedRange("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void GetMarkedRange___Should_throw_InvalidOperationException___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.GetMarkedRange(A.Dummy<string>()));
            var actual2 = Record.Exception(() => systemUnderTest2.GetMarkedRange(A.Dummy<string>()));

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("marker does not exist");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("marker does not exist");
        }

        [Fact]
        public static void GetMarkedRange___Should_return_marked_range___When_marker_exists_on_single_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.GetMarkedRange($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.GetCells().SingleOrDefault().ToCellReference().A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void GetMarkedRange___Should_return_marked_range___When_marker_exists_on_multiple_cells()
        {
            // Arrange
            var systemUnderTest1 = GetCursorWith3x3MarkedCanvas();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();
            systemUnderTest1.MoveToMarkedCell($"{DefaultMarkerNamePrefix}B2");
            systemUnderTest1.AddMarker(markerName);
            systemUnderTest1.MoveToMarkedCell($"{DefaultMarkerNamePrefix}A1");
            systemUnderTest1.AddMarker(markerName);

            systemUnderTest2.MoveToMarkedCell($"{DefaultMarkerNamePrefix}B2");
            systemUnderTest2.AddMarker(markerName);
            systemUnderTest2.MoveToMarkedCell($"{DefaultMarkerNamePrefix}A1");
            systemUnderTest2.AddMarker(markerName);
            systemUnderTest2.MoveToMarkedCell($"{DefaultMarkerNamePrefix}C2");
            systemUnderTest2.AddMarker(markerName);

            var expected1 = systemUnderTest1.Worksheet.GetRange(1, 2, 1, 2);
            var expected2 = systemUnderTest1.Worksheet.GetRange(1, 2, 1, 3);

            // Act
            var actual1 = systemUnderTest1.GetMarkedRange(markerName);
            var actual2 = systemUnderTest2.GetMarkedRange(markerName);

            // Assert
            actual1.RefersTo.Should().Be(expected1.RefersTo);
            actual2.RefersTo.Should().Be(expected2.RefersTo);
        }

        [Fact]
        public static void GetMarkedCells___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCells(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void GetMarkedCells___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCells("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void GetMarkedCells___Should_throw_InvalidOperationException___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.GetMarkedCells(A.Dummy<string>()));
            var actual2 = Record.Exception(() => systemUnderTest2.GetMarkedCells(A.Dummy<string>()));

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("marker does not exist");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("marker does not exist");
        }

        [Fact]
        public static void GetMarkedCells___Should_return_marked_cell___When_marker_exists_on_single_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.GetMarkedCells($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.SingleOrDefault().ToCellReference().A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void GetMarkedCells___Should_return_marked_cell___When_marker_exists_on_multiple_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();
            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveUp().MoveLeft();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveUp().MoveLeft();
            systemUnderTest.AddMarker(markerName);

            // Act
            var actual = systemUnderTest.GetMarkedCells(markerName);

            // Assert
            actual.Should().HaveCount(3);
            actual.Select(_ => _.ToCellReference().A1Reference).Should().BeEquivalentTo("A1", "B2", "C3");
        }

        [Fact]
        public static void GetMarkedCell___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCell(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void GetMarkedCell___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCell("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void GetMarkedCell___Should_throw_InvalidOperationException___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.GetMarkedCell(A.Dummy<string>()));
            var actual2 = Record.Exception(() => systemUnderTest2.GetMarkedCell(A.Dummy<string>()));

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("marker does not exist");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("marker does not exist");
        }

        [Fact]
        public static void GetMarkedCell___Should_throw_InvalidOperationException___When_marker_exists_on_multiple_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();
            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCell(markerName));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("Multiple cells are marked as");
        }

        [Fact]
        public static void GetMarkedCell___Should_return_marked_cell___When_marker_exists_on_single_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.GetMarkedCell($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.ToCellReference().A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void GetMarkedCellReference___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCellReference(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void GetMarkedCellReference___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCellReference("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void GetMarkedCellReference___Should_throw_InvalidOperationException___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.GetMarkedCellReference(A.Dummy<string>()));
            var actual2 = Record.Exception(() => systemUnderTest2.GetMarkedCellReference(A.Dummy<string>()));

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("marker does not exist");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("marker does not exist");
        }

        [Fact]
        public static void GetMarkedCellReference___Should_throw_InvalidOperationException___When_marker_exists_on_multiple_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();
            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);

            // Act
            var actual = Record.Exception(() => systemUnderTest.GetMarkedCellReference(markerName));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("Multiple cells are marked as");
        }

        [Fact]
        public static void GetMarkedCellReference___Should_return_marked_cell___When_marker_exists_on_single_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.GetMarkedCellReference($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.A1Reference.Should().Be("B2");
        }

        [Fact]
        public static void MoveToMarkedCell___Should_throw_ArgumentNullException___When_parameter_markerName_is_null()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveToMarkedCell(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("markerName");
        }

        [Fact]
        public static void MoveToMarkedCell___Should_throw_ArgumentException___When_parameter_markerName_is_white_space()
        {
            // Arrange
            var systemUnderTest = A.Dummy<CellCursor>();

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveToMarkedCell("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("markerName");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void MoveToMarkedCell___Should_throw_InvalidOperationException___When_marker_does_not_exist()
        {
            // Arrange
            var systemUnderTest1 = A.Dummy<CellCursor>();
            var systemUnderTest2 = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.MoveToMarkedCell(A.Dummy<string>()));
            var actual2 = Record.Exception(() => systemUnderTest2.MoveToMarkedCell(A.Dummy<string>()));

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("marker does not exist");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("marker does not exist");
        }

        [Fact]
        public static void MoveToMarkedCell___Should_throw_InvalidOperationException___When_marker_exists_on_multiple_cells()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            var markerName = A.Dummy<string>();
            systemUnderTest.Reset();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);
            systemUnderTest.MoveDown().MoveRight();
            systemUnderTest.AddMarker(markerName);

            // Act
            var actual = Record.Exception(() => systemUnderTest.MoveToMarkedCell(markerName));

            // Assert
            actual.Should().BeOfType<InvalidOperationException>();
            actual.Message.Should().Contain("Multiple cells are marked as");
        }

        [Fact]
        public static void MoveToMarkedCell___Should_move_cursor_to_marked_cell___When_marker_exists_on_single_cell()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            systemUnderTest.MoveToMarkedCell($"{DefaultMarkerNamePrefix}B3");

            // Assert
            systemUnderTest.RowNumber.Should().Be(3);
            systemUnderTest.ColumnNumber.Should().Be(2);

            systemUnderTest.StartRowNumber.Should().Be(1);
            systemUnderTest.StartColumnNumber.Should().Be(1);

            systemUnderTest.MaxRowNumber.Should().Be(3);
            systemUnderTest.MaxColumnNumber.Should().Be(3);
        }

        [Fact]
        public static void MoveToMarkedCell___Should_return_same_cursor___When_called()
        {
            // Arrange
            var systemUnderTest = GetCursorWith3x3MarkedCanvas();

            // Act
            var actual = systemUnderTest.MoveToMarkedCell($"{DefaultMarkerNamePrefix}B2");

            // Assert
            actual.Should().BeSameAs(systemUnderTest);
        }

        private static CellCursor GetCursorWithUniqueRowAndColumnNumbers()
        {
            var worksheet = A.Dummy<Worksheet>();

            var result = new CellCursor(worksheet, rowNumber: StartRowNumber, columnNumber: StartColumnNumber);
            result.MoveDown(MoveDownBy).MoveRight(MoveRightBy).MoveLeft(MoveLeftBy).MoveUp(MoveUpBy);

            var positions = new[] { result.RowNumber, result.ColumnNumber, result.StartRowNumber, result.StartColumnNumber, result.MaxRowNumber, result.MaxColumnNumber };
            positions.Distinct().Should().HaveCount(positions.Length);

            return result;
        }

        private static CellCursor GetCursorWith3x3MarkedCanvas(
            string markerNamePrefix = DefaultMarkerNamePrefix)
        {
            var worksheet = A.Dummy<Worksheet>();
            var result = new CellCursor(worksheet);

            result.AddMarker($"{markerNamePrefix}A1");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}B1");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}C1");

            result.ResetColumn();
            result.MoveDown();
            result.AddMarker($"{markerNamePrefix}A2");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}B2");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}C2");

            result.ResetColumn();
            result.MoveDown();
            result.AddMarker($"{markerNamePrefix}A3");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}B3");
            result.MoveRight();
            result.AddMarker($"{markerNamePrefix}C3");

            return result;
        }
    }
}
