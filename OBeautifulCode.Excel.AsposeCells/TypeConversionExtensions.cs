// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TypeConversionExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

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
            new { borderEdges }.Must().NotBeEqualTo(BorderEdges.Unknown);

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
            new { dataValidationKind }.Must().NotBeEqualTo(DataValidationKind.Unknown);

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
            new { dataValidationErrorAlertStyle }.Must().NotBeEqualTo(DataValidationErrorAlertStyle.Unknown);

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
            new { horizontalAlignment }.Must().NotBeEqualTo(HorizontalAlignment.Unknown);

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
            new { verticalAlignment }.Must().NotBeEqualTo(VerticalAlignment.Unknown);

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
    }
}
