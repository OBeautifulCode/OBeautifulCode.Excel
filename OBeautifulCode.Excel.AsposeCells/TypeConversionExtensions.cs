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
        /// <param name="dataValidationKind">The kind of validation.</param>
        /// <returns>
        /// A <see cref="ValidationType"/> converted from a <see cref="DataValidationKind"/>.
        /// </returns>
        public static ValidationType ToValidationType(
            this DataValidationKind dataValidationKind)
        {
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
        /// <param name="dataValidationOperator">The operator to apply to the data.</param>
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
        /// <param name="dataValidationErrorAlertStyle">The style of error alert to show on a data validation.</param>
        /// <returns>
        /// A <see cref="ValidationAlertType"/> converted from a <see cref="DataValidationErrorAlertStyle"/>.
        /// </returns>
        public static ValidationAlertType ToValidationAlertType(
            this DataValidationErrorAlertStyle dataValidationErrorAlertStyle)
        {
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
    }
}
