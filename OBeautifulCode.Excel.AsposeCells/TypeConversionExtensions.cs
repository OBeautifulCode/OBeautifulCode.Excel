// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TypeConversionExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using Aspose.Cells;
    using Aspose.Cells.Drawing;

    using static System.FormattableString;

    /// <summary>
    /// Contains extension methods to convert to/from OBeautifulCode.Excel and Aspose.Cells types.
    /// </summary>
    public static class TypeConversionExtensions
    {
        /// <summary>
        /// Converts a <see cref="BorderEdges"/> to a <see cref="BorderType"/>.
        /// </summary>
        /// <param name="borderEdges">The border edges to convert.</param>
        /// <returns>
        /// A <see cref="BorderType"/> converted from a <see cref="BorderEdges"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="borderEdges"/> is <see cref="BorderEdges.Unknown"/>.</exception>
        public static BorderType ToBorderType(
            this BorderEdges borderEdges)
        {
            if (borderEdges == BorderEdges.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(borderEdges)}' == '{BorderEdges.Unknown}'"), (Exception)null);
            }

            var result = default(BorderType);

            if (borderEdges.HasFlag(BorderEdges.BottomBorder))
            {
                result |= BorderType.BottomBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.DiagonalDown))
            {
                result |= BorderType.DiagonalDown;
            }

            if (borderEdges.HasFlag(BorderEdges.DiagonalUp))
            {
                result |= BorderType.DiagonalUp;
            }

            if (borderEdges.HasFlag(BorderEdges.LeftBorder))
            {
                result |= BorderType.LeftBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.RightBorder))
            {
                result |= BorderType.RightBorder;
            }

            if (borderEdges.HasFlag(BorderEdges.TopBorder))
            {
                result |= BorderType.TopBorder;
            }

            return result;
        }

        /// <summary>
        /// Converts a <see cref="BorderStyle"/> to a <see cref="CellBorderType"/>.
        /// </summary>
        /// <param name="borderStyle">The border style to convert.</param>
        /// <returns>
        /// A <see cref="CellBorderType"/> converted from a <see cref="BorderStyle"/>.
        /// </returns>
        public static CellBorderType ToCellBorderType(
            this BorderStyle borderStyle)
        {
            switch (borderStyle)
            {
                case BorderStyle.DashDot:
                    return CellBorderType.DashDot;
                case BorderStyle.DashDotDot:
                    return CellBorderType.DashDotDot;
                case BorderStyle.Dashed:
                    return CellBorderType.Dashed;
                case BorderStyle.Dotted:
                    return CellBorderType.Dotted;
                case BorderStyle.Double:
                    return CellBorderType.Double;
                case BorderStyle.Hair:
                    return CellBorderType.Hair;
                case BorderStyle.Medium:
                    return CellBorderType.Medium;
                case BorderStyle.MediumDashDot:
                    return CellBorderType.MediumDashDot;
                case BorderStyle.MediumDashDotDot:
                    return CellBorderType.MediumDashDotDot;
                case BorderStyle.MediumDashed:
                    return CellBorderType.MediumDashed;
                case BorderStyle.None:
                    return CellBorderType.None;
                case BorderStyle.SlantedDashDot:
                    return CellBorderType.SlantedDashDot;
                case BorderStyle.Thick:
                    return CellBorderType.Thick;
                case BorderStyle.Thin:
                    return CellBorderType.Thin;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(BorderStyle)} is not supported: {borderStyle}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="CommentBorderStyle"/> to a <see cref="ToMsoLineStyle"/>.
        /// </summary>
        /// <param name="commentBorderStyle">The comment border style to convert.</param>
        /// <returns>
        /// A <see cref="ToMsoLineStyle"/> converted from a <see cref="CommentBorderStyle"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="commentBorderStyle"/> is <see cref="CommentBorderStyle.Unknown"/>.</exception>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Mso", Justification = "This is the identifier used by Aspose.")]
        public static MsoLineStyle ToMsoLineStyle(
            this CommentBorderStyle commentBorderStyle)
        {
            if (commentBorderStyle == CommentBorderStyle.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(commentBorderStyle)}' == '{CommentBorderStyle.Unknown}'"), (Exception)null);
            }

            switch (commentBorderStyle)
            {
                case CommentBorderStyle.Single:
                    return MsoLineStyle.Single;
                case CommentBorderStyle.ThickBetweenThin:
                    return MsoLineStyle.ThickBetweenThin;
                case CommentBorderStyle.ThickThin:
                    return MsoLineStyle.ThickThin;
                case CommentBorderStyle.ThinThick:
                    return MsoLineStyle.ThinThick;
                case CommentBorderStyle.ThinThin:
                    return MsoLineStyle.ThinThin;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(CommentBorderStyle)} is not supported: {commentBorderStyle}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="DataValidationKind"/> to a <see cref="ValidationType"/>.
        /// </summary>
        /// <param name="dataValidationKind">The kind of data validation to convert.</param>
        /// <returns>
        /// A <see cref="ValidationType"/> converted from a <see cref="DataValidationKind"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="dataValidationKind"/> is <see cref="DataValidationKind.Unknown"/>.</exception>
        public static ValidationType ToValidationType(
            this DataValidationKind dataValidationKind)
        {
            if (dataValidationKind == DataValidationKind.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(dataValidationKind)}' == '{DataValidationKind.Unknown}'"), (Exception)null);
            }

            switch (dataValidationKind)
            {
                case DataValidationKind.AnyValue:
                    return ValidationType.AnyValue;
                case DataValidationKind.Custom:
                    return ValidationType.Custom;
                case DataValidationKind.Date:
                    return ValidationType.Date;
                case DataValidationKind.Decimal:
                    return ValidationType.Decimal;
                case DataValidationKind.List:
                    return ValidationType.List;
                case DataValidationKind.TextLength:
                    return ValidationType.TextLength;
                case DataValidationKind.Time:
                    return ValidationType.Time;
                case DataValidationKind.WholeNumber:
                    return ValidationType.WholeNumber;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(DataValidationKind)} is not supported: {dataValidationKind}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="DataValidationOperator"/> to a <see cref="OperatorType"/>.
        /// </summary>
        /// <param name="dataValidationOperator">The data validation operator to convert.</param>
        /// <returns>
        /// A <see cref="OperatorType"/> converted from a <see cref="DataValidationOperator"/>.
        /// </returns>
        public static OperatorType ToOperatorType(
            this DataValidationOperator dataValidationOperator)
        {
            switch (dataValidationOperator)
            {
                case DataValidationOperator.Between:
                    return OperatorType.Between;
                case DataValidationOperator.EqualTo:
                    return OperatorType.Equal;
                case DataValidationOperator.GreaterThan:
                    return OperatorType.GreaterThan;
                case DataValidationOperator.GreaterThanOrEqualTo:
                    return OperatorType.GreaterOrEqual;
                case DataValidationOperator.LessThan:
                    return OperatorType.LessThan;
                case DataValidationOperator.LessThanOrEqualTo:
                    return OperatorType.LessOrEqual;
                case DataValidationOperator.None:
                    return OperatorType.None;
                case DataValidationOperator.NotBetween:
                    return OperatorType.NotBetween;
                case DataValidationOperator.NotEqualTo:
                    return OperatorType.NotEqual;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(DataValidationOperator)} is not supported: {dataValidationOperator}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="DataValidationErrorAlertStyle"/> to a <see cref="ValidationAlertType"/>.
        /// </summary>
        /// <param name="dataValidationErrorAlertStyle">The data validation error alert style to convert.</param>
        /// <returns>
        /// A <see cref="ValidationAlertType"/> converted from a <see cref="DataValidationErrorAlertStyle"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="dataValidationErrorAlertStyle"/> is <see cref="DataValidationErrorAlertStyle.Unknown"/>.</exception>
        public static ValidationAlertType ToValidationAlertType(
            this DataValidationErrorAlertStyle dataValidationErrorAlertStyle)
        {
            if (dataValidationErrorAlertStyle == DataValidationErrorAlertStyle.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(dataValidationErrorAlertStyle)}' == '{DataValidationErrorAlertStyle.Unknown}'"), (Exception)null);
            }

            switch (dataValidationErrorAlertStyle)
            {
                case DataValidationErrorAlertStyle.Information:
                    return ValidationAlertType.Information;
                case DataValidationErrorAlertStyle.Stop:
                    return ValidationAlertType.Stop;
                case DataValidationErrorAlertStyle.Warning:
                    return ValidationAlertType.Warning;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(DataValidationErrorAlertStyle)} is not supported: {dataValidationErrorAlertStyle}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="HorizontalAlignment"/> to a <see cref="TextAlignmentType"/>.
        /// </summary>
        /// <param name="horizontalAlignment">The horizontal alignment to convert.</param>
        /// <returns>
        /// A <see cref="TextAlignmentType"/> converted from a <see cref="HorizontalAlignment"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="horizontalAlignment"/> is <see cref="HorizontalAlignment.Unknown"/>.</exception>
        public static TextAlignmentType ToTextAlignmentType(
            this HorizontalAlignment horizontalAlignment)
        {
            if (horizontalAlignment == HorizontalAlignment.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(horizontalAlignment)}' == '{HorizontalAlignment.Unknown}'"), (Exception)null);
            }

            switch (horizontalAlignment)
            {
                case HorizontalAlignment.Center:
                    return TextAlignmentType.Center;
                case HorizontalAlignment.Distributed:
                    return TextAlignmentType.Distributed;
                case HorizontalAlignment.Justify:
                    return TextAlignmentType.Justify;
                case HorizontalAlignment.Left:
                    return TextAlignmentType.Left;
                case HorizontalAlignment.Right:
                    return TextAlignmentType.Right;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(HorizontalAlignment)} is not supported: {horizontalAlignment}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="VerticalAlignment"/> to a <see cref="TextAlignmentType"/>.
        /// </summary>
        /// <param name="verticalAlignment">The vertical alignment to convert.</param>
        /// <returns>
        /// A <see cref="TextAlignmentType"/> converted from a <see cref="VerticalAlignment"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="verticalAlignment"/> is <see cref="VerticalAlignment.Unknown"/>.</exception>
        public static TextAlignmentType ToTextAlignmentType(
            this VerticalAlignment verticalAlignment)
        {
            if (verticalAlignment == VerticalAlignment.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(verticalAlignment)}' == '{VerticalAlignment.Unknown}'"), (Exception)null);
            }

            switch (verticalAlignment)
            {
                case VerticalAlignment.Bottom:
                    return TextAlignmentType.Bottom;
                case VerticalAlignment.Center:
                    return TextAlignmentType.Center;
                case VerticalAlignment.Distributed:
                    return TextAlignmentType.Distributed;
                case VerticalAlignment.Justify:
                    return TextAlignmentType.Justify;
                case VerticalAlignment.Top:
                    return TextAlignmentType.Top;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(VerticalAlignment)} is not supported: {verticalAlignment}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="UnderlineKind"/> to a <see cref="FontUnderlineType"/>.
        /// </summary>
        /// <param name="underlineKind">The kind of underline to convert.</param>
        /// <returns>
        /// A <see cref="FontUnderlineType"/> converted from a <see cref="UnderlineKind"/>.
        /// </returns>
        public static FontUnderlineType ToFontUnderlineType(
            this UnderlineKind underlineKind)
        {
            switch (underlineKind)
            {
                case UnderlineKind.Accounting:
                    return FontUnderlineType.Accounting;
                case UnderlineKind.Dash:
                    return FontUnderlineType.Dash;
                case UnderlineKind.DashDotDotHeavy:
                    return FontUnderlineType.DashDotDotHeavy;
                case UnderlineKind.DashDotHeavy:
                    return FontUnderlineType.DashDotHeavy;
                case UnderlineKind.DashedHeavy:
                    return FontUnderlineType.DashedHeavy;
                case UnderlineKind.DashLong:
                    return FontUnderlineType.DashLong;
                case UnderlineKind.DashLongHeavy:
                    return FontUnderlineType.DashLongHeavy;
                case UnderlineKind.DotDash:
                    return FontUnderlineType.DotDash;
                case UnderlineKind.DotDotDash:
                    return FontUnderlineType.DotDotDash;
                case UnderlineKind.Dotted:
                    return FontUnderlineType.Dotted;
                case UnderlineKind.DottedHeavy:
                    return FontUnderlineType.DottedHeavy;
                case UnderlineKind.Double:
                    return FontUnderlineType.Double;
                case UnderlineKind.DoubleAccounting:
                    return FontUnderlineType.DoubleAccounting;
                case UnderlineKind.Heavy:
                    return FontUnderlineType.Heavy;
                case UnderlineKind.None:
                    return FontUnderlineType.None;
                case UnderlineKind.Single:
                    return FontUnderlineType.Single;
                case UnderlineKind.Wave:
                    return FontUnderlineType.Wave;
                case UnderlineKind.WavyDouble:
                    return FontUnderlineType.WavyDouble;
                case UnderlineKind.WavyHeavy:
                    return FontUnderlineType.WavyHeavy;
                case UnderlineKind.Words:
                    return FontUnderlineType.Words;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(UnderlineKind)} is not supported: {underlineKind}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="ConditionalFormattingOperator"/> to a <see cref="OperatorType"/>.
        /// </summary>
        /// <param name="conditionalFormattingOperator">The conditional formatting operator to convert.</param>
        /// <returns>
        /// A <see cref="OperatorType"/> converted from a <see cref="ConditionalFormattingOperator"/>.
        /// </returns>
        public static OperatorType ToOperatorType(
            this ConditionalFormattingOperator conditionalFormattingOperator)
        {
            switch (conditionalFormattingOperator)
            {
                case ConditionalFormattingOperator.Between:
                    return OperatorType.Between;
                case ConditionalFormattingOperator.EqualTo:
                    return OperatorType.Equal;
                case ConditionalFormattingOperator.GreaterThan:
                    return OperatorType.GreaterThan;
                case ConditionalFormattingOperator.GreaterThanOrEqualTo:
                    return OperatorType.GreaterOrEqual;
                case ConditionalFormattingOperator.LessThan:
                    return OperatorType.LessThan;
                case ConditionalFormattingOperator.LessThanOrEqualTo:
                    return OperatorType.LessOrEqual;
                case ConditionalFormattingOperator.None:
                    return OperatorType.None;
                case ConditionalFormattingOperator.NotBetween:
                    return OperatorType.NotBetween;
                case ConditionalFormattingOperator.NotEqualTo:
                    return OperatorType.NotEqual;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(ConditionalFormattingOperator)} is not supported: {conditionalFormattingOperator}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="Format"/> to it's numeric value.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <returns>
        /// A numeric value representing a <see cref="Format"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="Format.Unknown"/>.</exception>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex.")]
        public static int ToFormatNumber(
            this Format format)
        {
            if (format == Format.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(format)}' == '{Format.Unknown}'"), (Exception)null);
            }

            switch (format)
            {
                case Format.General:
                    return 0;
                case Format.Decimal1:
                    return 1;
                case Format.Decimal2:
                    return 2;
                case Format.Decimal3:
                    return 3;
                case Format.Decimal4:
                    return 4;
                case Format.Currency1:
                    return 5;
                case Format.Currency2:
                    return 6;
                case Format.Currency3:
                    return 7;
                case Format.Currency4:
                    return 8;
                case Format.Percentage1:
                    return 9;
                case Format.Percentage2:
                    return 10;
                case Format.Scientific1:
                    return 11;
                case Format.Fraction1:
                    return 12;
                case Format.Fraction2:
                    return 13;
                case Format.Date1:
                    return 14;
                case Format.Date2:
                    return 15;
                case Format.Date3:
                    return 16;
                case Format.Date4:
                    return 17;
                case Format.Time1:
                    return 18;
                case Format.Time2:
                    return 19;
                case Format.Time3:
                    return 20;
                case Format.Time4:
                    return 21;
                case Format.Time5:
                    return 22;
                case Format.Accounting1:
                    return 37;
                case Format.Accounting2:
                    return 38;
                case Format.Accounting3:
                    return 39;
                case Format.Accounting4:
                    return 40;
                case Format.Accounting5:
                    return 41;
                case Format.Currency5:
                    return 42;
                case Format.Accounting6:
                    return 43;
                case Format.Currency6:
                    return 44;
                case Format.Time6:
                    return 45;
                case Format.Time7:
                    return 46;
                case Format.Time8:
                    return 47;
                case Format.Scientific2:
                    return 48;
                case Format.Text:
                    return 49;
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(Format)} is not supported: {format}"));
            }
        }

        /// <summary>
        /// Converts a <see cref="BuiltInDocumentPropertyKind"/> to a <see cref="string"/> key that can be used in a collection of built-in document properties.
        /// </summary>
        /// <param name="builtInDocumentPropertyKind">The built-in document property on a workbook.</param>
        /// <returns>
        /// A <see cref="string"/> key that can be used in a collection of built-in document properties, converted from a <see cref="BuiltInDocumentPropertyKind"/>.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="builtInDocumentPropertyKind"/> is <see cref="BuiltInDocumentPropertyKind.Unknown"/>.</exception>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex.")]
        public static string ToBuiltInDocumentPropertyCollectionKey(
            this BuiltInDocumentPropertyKind builtInDocumentPropertyKind)
        {
            if (builtInDocumentPropertyKind == BuiltInDocumentPropertyKind.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(builtInDocumentPropertyKind)}' == '{BuiltInDocumentPropertyKind.Unknown}'"), (Exception)null);
            }

            switch (builtInDocumentPropertyKind)
            {
                case BuiltInDocumentPropertyKind.Title:
                    return "Title";
                case BuiltInDocumentPropertyKind.Subject:
                    return "Subject";
                case BuiltInDocumentPropertyKind.Author:
                    return "Author";
                case BuiltInDocumentPropertyKind.Keywords:
                    return "Keywords";
                case BuiltInDocumentPropertyKind.Comments:
                    return "Comments";
                case BuiltInDocumentPropertyKind.Template:
                    return "Template";
                case BuiltInDocumentPropertyKind.LastAuthor:
                    return "Last Author";
                case BuiltInDocumentPropertyKind.RevisionNumber:
                    return "Revision Number";
                case BuiltInDocumentPropertyKind.ApplicationName:
                    return "Application Name";
                case BuiltInDocumentPropertyKind.LastPrintDate:
                    return "Last Print Date";
                case BuiltInDocumentPropertyKind.CreationDate:
                    return "Creation Date";
                case BuiltInDocumentPropertyKind.LastSaveTime:
                    return "Last Save Time";
                case BuiltInDocumentPropertyKind.TotalEditingTime:
                    return "Total Editing Time";
                case BuiltInDocumentPropertyKind.NumberOfPages:
                    return "Number of Pages";
                case BuiltInDocumentPropertyKind.NumberOfWords:
                    return "Number of Words";
                case BuiltInDocumentPropertyKind.NumberOfCharacters:
                    return "Number of Characters";
                case BuiltInDocumentPropertyKind.Security:
                    return "Security";
                case BuiltInDocumentPropertyKind.Category:
                    return "Category";
                case BuiltInDocumentPropertyKind.Format:
                    return "Format";
                case BuiltInDocumentPropertyKind.Manager:
                    return "Manager";
                case BuiltInDocumentPropertyKind.Company:
                    return "Company";
                case BuiltInDocumentPropertyKind.NumberOfBytes:
                    return "Number of Bytes";
                case BuiltInDocumentPropertyKind.NumberOfLines:
                    return "Number of Lines";
                case BuiltInDocumentPropertyKind.NumberOfParagraphs:
                    return "Number of Paragraphs";
                case BuiltInDocumentPropertyKind.NumberOfSlides:
                    return "Number of Slides";
                case BuiltInDocumentPropertyKind.NumberOfNotes:
                    return "Number of Notes";
                case BuiltInDocumentPropertyKind.NumberOfHiddenSlides:
                    return "Number of Hidden Slides";
                case BuiltInDocumentPropertyKind.NumberOfMultimediaClips:
                    return "Number of Multimedia Clips";
                default:
                    throw new NotSupportedException(Invariant($"This {nameof(BuiltInDocumentPropertyKind)} is not supported: {builtInDocumentPropertyKind}"));
            }
        }
    }
}
