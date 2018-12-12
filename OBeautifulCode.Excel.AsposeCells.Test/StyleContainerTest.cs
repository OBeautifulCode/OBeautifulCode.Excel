// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StyleContainerTest.cs" company="OBeautifulCode">
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

    public static class StyleContainerTest
    {
        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_style_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new StyleContainer(null, new StyleFlag()));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("style");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_styleFlag_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new StyleContainer(new Style(), null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("styleFlag");
        }

        [Fact]
        public static void Style___Should_return_same_style_passed_to_constructor___When_getting()
        {
            // Arrange
            var style = new Style();
            var styleFlag = new StyleFlag();
            var systemUnderTest = new StyleContainer(style, styleFlag);

            // Act
            var actual = systemUnderTest.Style;

            // Assert
            actual.Should().Be(style);
        }

        [Fact]
        public static void StyleFlag___Should_return_same_styleFlag_passed_to_constructor___When_getting()
        {
            // Arrange
            var style = new Style();
            var styleFlag = new StyleFlag();
            var systemUnderTest = new StyleContainer(style, styleFlag);

            // Act
            var actual = systemUnderTest.StyleFlag;

            // Assert
            actual.Should().Be(styleFlag);
        }

        [Fact]
        public static void BuildNew___Should_throw_ArgumentNullException___When_parameter_workbook_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => StyleContainer.BuildNew(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("workbook");
        }

        [Fact]
        public static void BuildNew___Should_return_a_style_container___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();

            // Act
            var actual = StyleContainer.BuildNew(worksheet.Workbook);

            // Assert
            actual.Should().NotBeNull();
            actual.Style.Should().NotBeNull();
            actual.StyleFlag.Should().NotBeNull();
        }

        [Fact(Skip = "No good way to test this.")]
        public static void BuildNew___Should_build_new_style_container_whose_underlying_Style_is_registered_with_the_workbook___When_called()
        {
        }

        [Fact]
        public static void BuildExistingFromCell___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => StyleContainer.BuildUsingExistingCellStyle(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void BuildUsingExistingCellStyle___Should_build_style_container_using_existing_cell_style___When_called()
        {
            // Arrange
            var worksheet = A.Dummy<Worksheet>();
            var cell = worksheet.Cells[1, 1];

            // Act
            var actual = StyleContainer.BuildUsingExistingCellStyle(cell);

            // Assert
            actual.Style.Should().Be(cell.GetStyle());
        }

        [Fact]
        public static void ApplyToRange___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange
            var styleContainer = new StyleContainer(new Style(), new StyleFlag());

            // Act
            var actual = Record.Exception(() => styleContainer.ApplyToRange(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("range");
        }

        [Fact(Skip = "Too hard to test.")]
        public static void ApplyToRange___Should_apply_the_style_container_to_the_range___When_called()
        {
        }

        [Fact]
        public static void ApplyToCell___Should_throw_ArgumentNullException___When_parameter_range_is_null()
        {
            // Arrange
            var styleContainer = new StyleContainer(new Style(), new StyleFlag());

            // Act
            var actual = Record.Exception(() => styleContainer.ApplyToCell(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact(Skip = "Too hard to test.")]
        public static void ApplyToCell___Should_apply_the_style_container_to_the_cell___When_called()
        {
        }
    }
}
