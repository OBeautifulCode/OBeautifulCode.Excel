// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetExtensionsTest.Read.cs" company="OBeautifulCode">
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

    public static partial class WorksheetExtensionsTest
    {
        [Fact]
        public static void GetRange___Should_throw_ArgumentNullException___When_parameter_worksheet_is_null()
        {
            // Arrange
            var startRowNumber = 5;
            var endRowNumber = 10;
            var startColumnNumber = 30;
            var endColumnNumber = 35;

            // Act
            var actual = Record.Exception(() => WorksheetExtensions.GetRange(null, startRowNumber, endRowNumber, startColumnNumber, endColumnNumber));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheet");
        }

        [Fact]
        public static void GetRange___Should_throw_ArgumentOutOfRangeException___When_parameter_startRowNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var endRowNumber = 10;
            var startColumnNumber = 30;
            var endColumnNumber = 35;

            // Act
            var actual1 = Record.Exception(() => worksheet.GetRange(0, endRowNumber, startColumnNumber, endColumnNumber));
            var actual2 = Record.Exception(() => worksheet.GetRange(A.Dummy<NegativeInteger>(), endRowNumber, startColumnNumber, endColumnNumber));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("startRowNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("startRowNumber");
        }

        [Fact]
        public static void GetRange___Should_throw_ArgumentOutOfRangeException___When_parameter_startColumnNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var startRowNumber = 5;
            var endRowNumber = 10;
            var endColumnNumber = 35;

            // Act
            var actual1 = Record.Exception(() => worksheet.GetRange(startRowNumber, endRowNumber, 0, endColumnNumber));
            var actual2 = Record.Exception(() => worksheet.GetRange(startRowNumber, endRowNumber, A.Dummy<NegativeInteger>(), endColumnNumber));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("startColumnNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("startColumnNumber");
        }

        [Fact]
        public static void GetRange___Should_throw_ArgumentOutOfRangeException___When_parameter_endRowNumber_is_less_than_startRowNumber()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var startRowNumber = 5;
            var startColumnNumber = 30;
            var endColumnNumber = 35;

            // Act
            var actual1 = Record.Exception(() => worksheet.GetRange(startRowNumber, startRowNumber - 1, startColumnNumber, endColumnNumber));
            var actual2 = Record.Exception(() => worksheet.GetRange(startRowNumber, A.Dummy<PositiveInteger>().ThatIs(_ => _ < startRowNumber, -1), startColumnNumber, endColumnNumber));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("endRowNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("endRowNumber");
        }

        [Fact]
        public static void GetRange___Should_throw_ArgumentOutOfRangeException___When_parameter_endColumnNumber_is_less_than_startRowNumber()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var startRowNumber = 5;
            var endRowNumber = 10;
            var startColumnNumber = 30;

            // Act
            var actual1 = Record.Exception(() => worksheet.GetRange(startRowNumber, endRowNumber, startColumnNumber, startColumnNumber - 1));
            var actual2 = Record.Exception(() => worksheet.GetRange(startRowNumber, endRowNumber, startColumnNumber, A.Dummy<PositiveInteger>().ThatIs(_ => _ < startColumnNumber)));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("endColumnNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("endColumnNumber");
        }

        [Fact]
        public static void GetRange___Should_return_range___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual1 = worksheet.GetRange(2, 2, 2, 2);
            var actual2 = worksheet.GetRange(2, 4, 2, 2);
            var actual3 = worksheet.GetRange(2, 2, 2, 4);
            var actual4 = worksheet.GetRange(2, 4, 2, 4);

            // Assert
            actual1.RefersTo.Should().Be("=Sheet1!$B$2");
            actual2.RefersTo.Should().Be("=Sheet1!$B$2:$B$4");
            actual3.RefersTo.Should().Be("=Sheet1!$B$2:$D$2");
            actual4.RefersTo.Should().Be("=Sheet1!$B$2:$D$4");
        }

        [Fact]
        public static void GetCell___Should_throw_ArgumentNullException___When_parameter_worksheet_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => WorksheetExtensions.GetCell(null, A.Dummy<PositiveInteger>(), A.Dummy<PositiveInteger>()));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheet");
        }

        [Fact]
        public static void GetCell___Should_throw_ArgumentOutOfRangeException___When_parameter_rowNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual1 = Record.Exception(() => worksheet.GetCell(0, A.Dummy<PositiveInteger>()));
            var actual2 = Record.Exception(() => worksheet.GetCell(A.Dummy<NegativeInteger>(), A.Dummy<PositiveInteger>()));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("rowNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("rowNumber");
        }

        [Fact]
        public static void GetCell___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_less_than_1()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual1 = Record.Exception(() => worksheet.GetCell(A.Dummy<PositiveInteger>(), 0));
            var actual2 = Record.Exception(() => worksheet.GetCell(A.Dummy<PositiveInteger>(), A.Dummy<NegativeInteger>()));

            // Assert
            actual1.Should().BeOfType<ArgumentOutOfRangeException>();
            actual1.Message.Should().Contain("columnNumber");

            actual2.Should().BeOfType<ArgumentOutOfRangeException>();
            actual2.Message.Should().Contain("columnNumber");
        }

        [Fact]
        public static void GetCell___Should_return_cell___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual = worksheet.GetCell(3, 4);

            // Assert
            actual.ToCellReference().A1Reference.Should().Be("D3");
        }
    }
}
