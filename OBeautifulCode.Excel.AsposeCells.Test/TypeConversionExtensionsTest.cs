// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TypeConversionExtensionsTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;

    using Aspose.Cells;
    using Aspose.Cells.Drawing;

    using FluentAssertions;

    using OBeautifulCode.Enum.Recipes;

    using Xunit;

    public static class TypeConversionExtensionsTest
    {
        [Fact]
        public static void ToBorderType__Should_throw_ArgumentOutOfRangeException___When_parameter_borderEdges_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => BorderEdges.Unknown.ToBorderType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(BorderEdges.Unknown));
        }

        [Fact]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flag", Justification = "This is the most descriptive term to use in this case.")]
        public static void ToBorderType__Should_convert_borderEdges_to_a_BorderType___When_borderEdges_has_single_flag()
        {
            // Arrange
            var flags = EnumExtensions.GetIndividualFlags<BorderEdges>().Skip(1).ToList();
            var expected = flags.Select(_ => Enum.Parse(typeof(BorderType), _.ToString())).ToList();

            // Act
            var actual = flags.Select(_ => _.ToBorderType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Justification = "This is the most descriptive term to use in this case.")]
        public static void ToBorderType__Should_convert_borderEdges_to_a_BorderType___When_borderEdges_has_multiple_flags()
        {
            // Arrange
            var borderEdges1 = BorderEdges.DiagonalDown | BorderEdges.TopBorder;
            var borderEdges2 = BorderEdges.Outline;

            var expected1 = BorderType.DiagonalDown | BorderType.TopBorder;
            var expected2 = BorderType.BottomBorder | BorderType.TopBorder | BorderType.LeftBorder | BorderType.RightBorder;

            // Act
            var actual1 = borderEdges1.ToBorderType();
            var actual2 = borderEdges2.ToBorderType();

            // Act
            actual1.Should().Be(expected1);
            actual2.Should().Be(expected2);
        }

        [Fact]
        public static void ToCellBorderType__Should_convert_borderStyle_to_a_CellBorderType___When_called()
        {
            // Arrange
            var borderStyles = EnumExtensions.GetEnumValues<BorderStyle>();
            var expected = borderStyles.Select(_ => Enum.Parse(typeof(CellBorderType), _.ToString())).ToList();

            // Act
            var actual = borderStyles.Select(_ => _.ToCellBorderType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Mso", Justification = "This is the identifier used by Aspose.")]
        public static void ToMsoLineStyle__Should_throw_ArgumentOutOfRangeException___When_parameter_commentBorderStyle_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CommentBorderStyle.Unknown.ToMsoLineStyle());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(CommentBorderStyle.Unknown));
        }

        [Fact]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Mso", Justification = "This is the identifier used by Aspose.")]
        public static void ToMsoLineStyle__Should_convert_commentBorderStyle_to_a_MsoLineStyle___When_called()
        {
            // Arrange
            var commentBorderStyles = EnumExtensions.GetEnumValues<CommentBorderStyle>().Where(_ => _ != CommentBorderStyle.Unknown).ToList();
            var expected = commentBorderStyles.Select(_ => Enum.Parse(typeof(MsoLineStyle), _.ToString())).ToList();

            // Act
            var actual = commentBorderStyles.Select(_ => _.ToMsoLineStyle()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToValidationType__Should_throw_ArgumentOutOfRangeException___When_parameter_dataValidationKind_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => DataValidationKind.Unknown.ToValidationType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(DataValidationKind.Unknown));
        }

        [Fact]
        public static void ToValidationType__Should_convert_dataValidationKind_to_a_ValidationType___When_called()
        {
            // Arrange
            var dataValidationKinds = EnumExtensions.GetEnumValues<DataValidationKind>().Where(_ => _ != DataValidationKind.Unknown).ToList();
            var expected = dataValidationKinds.Select(_ => Enum.Parse(typeof(ValidationType), _.ToString())).ToList();

            // Act
            var actual = dataValidationKinds.Select(_ => _.ToValidationType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToOperatorType__Should_convert_dataValidationOperator_to_an_OperatorType___When_called()
        {
            // Arrange
            var dataValidationOperators = new[]
            {
                DataValidationOperator.Between,
                DataValidationOperator.EqualTo,
                DataValidationOperator.GreaterThan,
                DataValidationOperator.GreaterThanOrEqualTo,
                DataValidationOperator.LessThan,
                DataValidationOperator.LessThanOrEqualTo,
                DataValidationOperator.None,
                DataValidationOperator.NotBetween,
                DataValidationOperator.NotEqualTo,
            };

            var expected = new[]
            {
                OperatorType.Between,
                OperatorType.Equal,
                OperatorType.GreaterThan,
                OperatorType.GreaterOrEqual,
                OperatorType.LessThan,
                OperatorType.LessOrEqual,
                OperatorType.None,
                OperatorType.NotBetween,
                OperatorType.NotEqual,
            };

            // Act
            var actual = dataValidationOperators.Select(_ => _.ToOperatorType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToValidationAlertType__Should_throw_ArgumentOutOfRangeException___When_parameter_dataValidationErrorAlertStyle_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => DataValidationErrorAlertStyle.Unknown.ToValidationAlertType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(DataValidationErrorAlertStyle.Unknown));
        }

        [Fact]
        public static void ToValidationAlertType__Should_convert_dataValidationErrorAlertStyle_to_a_ValidationAlertType___When_called()
        {
            // Arrange
            var dataValidationErrorAlertStyles = EnumExtensions.GetEnumValues<DataValidationErrorAlertStyle>().Where(_ => _ != DataValidationErrorAlertStyle.Unknown).ToList();
            var expected = dataValidationErrorAlertStyles.Select(_ => Enum.Parse(typeof(ValidationAlertType), _.ToString())).ToList();

            // Act
            var actual = dataValidationErrorAlertStyles.Select(_ => _.ToValidationAlertType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToTextAlignmentType__Should_throw_ArgumentOutOfRangeException___When_parameter_horizontalAlignment_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => HorizontalAlignment.Unknown.ToTextAlignmentType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(HorizontalAlignment.Unknown));
        }

        [Fact]
        public static void ToTextAlignmentType__Should_convert_horizontalAlignment_to_a_TextAlignmentType___When_called()
        {
            // Arrange
            var horizontalAlignments = EnumExtensions.GetEnumValues<HorizontalAlignment>().Where(_ => _ != HorizontalAlignment.Unknown).ToList();
            var expected = horizontalAlignments.Select(_ => Enum.Parse(typeof(TextAlignmentType), _.ToString())).ToList();

            // Act
            var actual = horizontalAlignments.Select(_ => _.ToTextAlignmentType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToTextAlignmentType__Should_throw_ArgumentOutOfRangeException___When_parameter_verticalAlignment_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => VerticalAlignment.Unknown.ToTextAlignmentType());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(VerticalAlignment.Unknown));
        }

        [Fact]
        public static void ToTextAlignmentType__Should_convert_verticalAlignment_to_a_TextAlignmentType___When_called()
        {
            // Arrange
            var verticalAlignments = EnumExtensions.GetEnumValues<VerticalAlignment>().Where(_ => _ != VerticalAlignment.Unknown).ToList();
            var expected = verticalAlignments.Select(_ => Enum.Parse(typeof(TextAlignmentType), _.ToString())).ToList();

            // Act
            var actual = verticalAlignments.Select(_ => _.ToTextAlignmentType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToFontUnderlineType__Should_convert_underlineKind_to_a_FontUnderlineType___When_called()
        {
            // Arrange
            var underlineKinds = EnumExtensions.GetEnumValues<UnderlineKind>().ToList();
            var expected = underlineKinds.Select(_ => Enum.Parse(typeof(FontUnderlineType), _.ToString())).ToList();

            // Act
            var actual = underlineKinds.Select(_ => _.ToFontUnderlineType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToOperatorType__Should_convert_conditionalFormattingOperator_to_an_OperatorType___When_called()
        {
            // Arrange
            var conditionalFormattingOperators = new[]
            {
                ConditionalFormattingOperator.Between,
                ConditionalFormattingOperator.EqualTo,
                ConditionalFormattingOperator.GreaterThan,
                ConditionalFormattingOperator.GreaterThanOrEqualTo,
                ConditionalFormattingOperator.LessThan,
                ConditionalFormattingOperator.LessThanOrEqualTo,
                ConditionalFormattingOperator.None,
                ConditionalFormattingOperator.NotBetween,
                ConditionalFormattingOperator.NotEqualTo,
            };

            var expected = new[]
            {
                OperatorType.Between,
                OperatorType.Equal,
                OperatorType.GreaterThan,
                OperatorType.GreaterOrEqual,
                OperatorType.LessThan,
                OperatorType.LessOrEqual,
                OperatorType.None,
                OperatorType.NotBetween,
                OperatorType.NotEqual,
            };

            // Act
            var actual = conditionalFormattingOperators.Select(_ => _.ToOperatorType()).ToList();

            // Act
            actual.Should().Equal(expected);
        }

        [Fact]
        public static void ToFormatNumber__Should_throw_ArgumentOutOfRangeException___When_parameter_format_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => Format.Unknown.ToFormatNumber());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(Format.Unknown));
        }

        [Fact(Skip = "No great way to test get other than just repeating the mapping in the method itself.")]
        public static void ToFormatNumber__Should_convert_format_to_its_numeric_representation___When_called()
        {
        }

        [Fact]
        public static void ToBuiltInDocumentPropertyCollectionKey__Should_throw_ArgumentOutOfRangeException___When_parameter_builtInDocumentPropertyKind_is_Unknown()
        {
            // Arrange, Act
            var actual = Record.Exception(() => BuiltInDocumentPropertyKind.Unknown.ToBuiltInDocumentPropertyCollectionKey());

            // Assert
            actual.Should().BeOfType<ArgumentOutOfRangeException>();
            actual.Message.Should().Contain(nameof(BuiltInDocumentPropertyKind.Unknown));
        }

        [Fact]
        public static void ToBuiltInDocumentPropertyCollectionKey__Should_convert_builtInDocumentPropertyKind_to_an_string_that_can_be_used_as_a_key_in_a_collection_of_built_in_document_properties___When_called()
        {
            // Arrange
            var builtInDocumentPropertyKinds = new[]
            {
                BuiltInDocumentPropertyKind.Title,
                BuiltInDocumentPropertyKind.Subject,
                BuiltInDocumentPropertyKind.Author,
                BuiltInDocumentPropertyKind.Keywords,
                BuiltInDocumentPropertyKind.Comments,
                BuiltInDocumentPropertyKind.Template,
                BuiltInDocumentPropertyKind.LastAuthor,
                BuiltInDocumentPropertyKind.RevisionNumber,
                BuiltInDocumentPropertyKind.ApplicationName,
                BuiltInDocumentPropertyKind.LastPrintDate,
                BuiltInDocumentPropertyKind.CreationDate,
                BuiltInDocumentPropertyKind.LastSaveTime,
                BuiltInDocumentPropertyKind.TotalEditingTime,
                BuiltInDocumentPropertyKind.NumberOfPages,
                BuiltInDocumentPropertyKind.NumberOfWords,
                BuiltInDocumentPropertyKind.NumberOfCharacters,
                BuiltInDocumentPropertyKind.Security,
                BuiltInDocumentPropertyKind.Category,
                BuiltInDocumentPropertyKind.Format,
                BuiltInDocumentPropertyKind.Manager,
                BuiltInDocumentPropertyKind.Company,
                BuiltInDocumentPropertyKind.NumberOfBytes,
                BuiltInDocumentPropertyKind.NumberOfLines,
                BuiltInDocumentPropertyKind.NumberOfParagraphs,
                BuiltInDocumentPropertyKind.NumberOfSlides,
                BuiltInDocumentPropertyKind.NumberOfNotes,
                BuiltInDocumentPropertyKind.NumberOfHiddenSlides,
                BuiltInDocumentPropertyKind.NumberOfMultimediaClips,
            };

            var expected = new[]
            {
                "Title",
                "Subject",
                "Author",
                "Keywords",
                "Comments",
                "Template",
                "Last Author",
                "Revision Number",
                "Application Name",
                "Last Print Date",
                "Creation Date",
                "Last Save Time",
                "Total Editing Time",
                "Number of Pages",
                "Number of Words",
                "Number of Characters",
                "Security",
                "Category",
                "Format",
                "Manager",
                "Company",
                "Number of Bytes",
                "Number of Lines",
                "Number of Paragraphs",
                "Number of Slides",
                "Number of Notes",
                "Number of Hidden Slides",
                "Number of Multimedia Clips",
            };

            // Act
            var actual = builtInDocumentPropertyKinds.Select(_ => _.ToBuiltInDocumentPropertyCollectionKey()).ToList();

            // Act
            actual.Should().Equal(expected);
        }
    }
}
