// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeStyleExtensionsTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System.Drawing;

    using FakeItEasy;

    using FluentAssertions;

    using OBeautifulCode.AutoFakeItEasy;

    using Xunit;

    public static class RangeStyleExtensionsTest
    {
        [Fact]
        public static void CloneWithFontColor__Should_return_equivalent_but_not_the_same_RangeStyle___When_fontColor_is_equal_to_rangeStyle_FontColor()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithFontColor(expected.FontColor);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithFontColor__Should_return_different_RangeStyle_with_updated_FontColor___When_fontColor_is_not_equal_to_rangeStyle_FontColor()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var fontColor = A.Dummy<Color?>().ThatIsNot(expected.FontColor);

            // Act
            var actual = expected.CloneWithFontColor(fontColor);

            // Assert
            actual.Should().NotBe(expected);
            actual.FontColor.Should().Be(fontColor);
        }

        [Fact]
        public static void CloneWithFontSize__Should_return_equivalent_but_not_the_same_RangeStyle___When_fontSize_is_equal_to_rangeStyle_FontSize()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithFontSize(expected.FontSize);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithFontSize__Should_return_different_RangeStyle_with_updated_FontSize___When_fontSize_is_not_equal_to_rangeStyle_FontSize()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var fontSize = A.Dummy<int?>().ThatIsNot(expected.FontSize);

            // Act
            var actual = expected.CloneWithFontSize(fontSize);

            // Assert
            actual.Should().NotBe(expected);
            actual.FontSize.Should().Be(fontSize);
        }

        [Fact]
        public static void CloneWithFontIsBold__Should_return_equivalent_but_not_the_same_RangeStyle___When_fontIsBold_is_equal_to_rangeStyle_FontIsBold()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithFontIsBold(expected.FontIsBold);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithFontIsBold__Should_return_different_RangeStyle_with_updated_FontIsBold___When_fontIsBold_is_not_equal_to_rangeStyle_FontIsBold()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var fontIsBold = A.Dummy<bool?>().ThatIsNot(expected.FontIsBold);

            // Act
            var actual = expected.CloneWithFontIsBold(fontIsBold);

            // Assert
            actual.Should().NotBe(expected);
            actual.FontIsBold.Should().Be(fontIsBold);
        }

        [Fact]
        public static void CloneWithBackgroundColor__Should_return_equivalent_but_not_the_same_RangeStyle___When_backgroundColor_is_equal_to_rangeStyle_BackgroundColor()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithBackgroundColor(expected.BackgroundColor);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithBackgroundColor__Should_return_different_RangeStyle_with_updated_BackgroundColor___When_backgroundColor_is_not_equal_to_rangeStyle_BackgroundColor()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var backgroundColor = A.Dummy<Color?>().ThatIsNot(expected.BackgroundColor);

            // Act
            var actual = expected.CloneWithBackgroundColor(backgroundColor);

            // Assert
            actual.Should().NotBe(expected);
            actual.BackgroundColor.Should().Be(backgroundColor);
        }

        [Fact]
        public static void CloneWithRowHeightInPixels__Should_return_equivalent_but_not_the_same_RangeStyle___When_rowHeightInPixels_is_equal_to_rangeStyle_RowHeightInPixels()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithRowHeightInPixels(expected.RowHeightInPixels);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithRowHeightInPixels__Should_return_different_RangeStyle_with_updated_RowHeightInPixels___When_rowHeightInPixels_is_not_equal_to_rangeStyle_RowHeightInPixels()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var rowHeightInPixels = A.Dummy<int?>().ThatIsNot(expected.RowHeightInPixels);

            // Act
            var actual = expected.CloneWithRowHeightInPixels(rowHeightInPixels);

            // Assert
            actual.Should().NotBe(expected);
            actual.RowHeightInPixels.Should().Be(rowHeightInPixels);
        }

        [Fact]
        public static void CloneWithColumnWidthInPixels__Should_return_equivalent_but_not_the_same_RangeStyle___When_columnWidthInPixels_is_equal_to_rangeStyle_ColumnWidthInPixels()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();

            // Act
            var actual = expected.CloneWithColumnWidthInPixels(expected.ColumnWidthInPixels);

            // Assert
            actual.Should().Be(expected);
            actual.Should().NotBeSameAs(expected);
        }

        [Fact]
        public static void CloneWithColumnWidthInPixels__Should_return_different_RangeStyle_with_updated_ColumnWidthInPixels___When_columnWidthInPixels_is_not_equal_to_rangeStyle_ColumnWidthInPixels()
        {
            // Arrange
            var expected = A.Dummy<RangeStyle>();
            var columnWidthInPixels = A.Dummy<int?>().ThatIsNot(expected.ColumnWidthInPixels);

            // Act
            var actual = expected.CloneWithColumnWidthInPixels(columnWidthInPixels);

            // Assert
            actual.Should().NotBe(expected);
            actual.ColumnWidthInPixels.Should().Be(columnWidthInPixels);
        }
    }
}
