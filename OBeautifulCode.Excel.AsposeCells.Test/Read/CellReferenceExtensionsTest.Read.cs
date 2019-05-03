// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReferenceExtensionsTest.Read.cs" company="OBeautifulCode">
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

    public static partial class CellReferenceExtensionsTest
    {
        [Fact]
        public static void GetCell___Should_throw_ArgumentNullException___When_parameter_cellReference_is_null()
        {
            // Arrange
            var workbook = A.Dummy<Worksheet>().Workbook;

            // Act
            var actual = Record.Exception(() => CellReferenceExtensions.GetCell(null, workbook));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cellReference");
        }

        [Fact]
        public static void GetCell___Should_throw_ArgumentNullException___When_parameter_workbook_is_null()
        {
            // Arrange
            var cellReference = A.Dummy<CellReference>();

            // Act
            var actual = Record.Exception(() => cellReference.GetCell(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("workbook");
        }

        [Fact]
        public static void GetCell___Should_return_cell_corresponding_to_cellReference___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            var expected = worksheet.Cells["C9"];

            var systemUnderTest = expected.ToCellReference();

            // Act
            var actual = systemUnderTest.GetCell(worksheet.Workbook);

            // Assert
            actual.Should().Be(expected);
        }
    }
}
