// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeStyleExtensions.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.Linq;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using Border = OBeautifulCode.Excel.Border;
    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Style-related extensions methods on type <see cref="Range"/>.
    /// </summary>
    public static class RangeStyleExtensions
    {
        /// <summary>
        /// Sets the background color of a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="color">The color.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetBackgroundColor(
            this Range range,
            Color? color,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (color != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Pattern = BackgroundType.Solid;
                    _.Style.ForegroundColor = (Color)color;
                    _.StyleFlag.CellShading = true;
                });
            }
        }

        /// <summary>
        /// Sets the font color for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="color">The color.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontColor(
            this Range range,
            Color? color,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (color != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.Color = (Color)color;
                    _.StyleFlag.FontColor = true;
                });
            }
        }

        /// <summary>
        /// Sets the font name for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="fontName">The font name.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontName(
            this Range range,
            string fontName,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (fontName != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.Name = fontName;
                    _.StyleFlag.FontName = true;
                });
            }
        }

        /// <summary>
        /// Sets the font size for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="fontSize">The font size.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontSize(
            this Range range,
            int? fontSize,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (fontSize != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.Size = (int)fontSize;
                    _.StyleFlag.FontSize = true;
                });
            }
        }

        /// <summary>
        /// Sets italics on the font.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="fontIsItalic">Determines whether the font is italic.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontIsItalic(
            this Range range,
            bool? fontIsItalic,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (fontIsItalic != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.IsItalic = (bool)fontIsItalic;
                    _.StyleFlag.FontItalic = true;
                });
            }
        }

        /// <summary>
        /// Sets bold on the font.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="fontIsBold">Determines whether the font is bold.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontIsBold(
            this Range range,
            bool? fontIsBold,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (fontIsBold != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.IsBold = (bool)fontIsBold;
                    _.StyleFlag.FontBold = true;
                });
            }
        }

        /// <summary>
        /// Sets underline on the font.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="underlineKind">The kind of underline.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontUnderline(
            this Range range,
            UnderlineKind? underlineKind,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (underlineKind != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Font.Underline = ((UnderlineKind)underlineKind).ToFontUnderlineType();
                    _.StyleFlag.FontUnderline = true;
                });
            }
        }

        /// <summary>
        /// Sets the rotation angle for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="rotationAngle">The rotation angle.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFontRotationAngle(
            this Range range,
            int? rotationAngle,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (rotationAngle != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.RotationAngle = (int)rotationAngle;
                    _.StyleFlag.Rotation = true;
                });
            }
        }

        /// <summary>
        /// Sets text wrapping.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="textIsWrapped">Determines whether text is wrapped.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetTextIsWrapped(
            this Range range,
            bool? textIsWrapped,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (textIsWrapped != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.IsTextWrapped = (bool)textIsWrapped;
                    _.StyleFlag.WrapText = true;
                });
            }
        }

        /// <summary>
        /// Sets the text format.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="format">The format.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetFormat(
            this Range range,
            Format? format,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (format != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Number = ((Format)format).ToFormatNumber();
                    _.StyleFlag.NumberFormat = true;
                });
            }
        }

        /// <summary>
        /// Sets a custom format.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="customFormatString">The custom string to use.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string", Justification = "This is a good usage of 'string'.")]
        public static void SetCustomFormat(
            this Range range,
            string customFormatString,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (customFormatString != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.Custom = customFormatString;
                    _.StyleFlag.NumberFormat = true;
                });
            }
        }

        /// <summary>
        /// Sets the indent level for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="indentLevel">The indent level.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetIndentLevel(
            this Range range,
            int? indentLevel,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (indentLevel != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.IndentLevel = (int)indentLevel;
                    _.StyleFlag.Indent = true;
                });
            }
        }

        /// <summary>
        /// Sets the vertical alignment of a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="verticalAlignment">The vertical alignment.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetVerticalAlignment(
            this Range range,
            VerticalAlignment? verticalAlignment,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (verticalAlignment != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.VerticalAlignment = ((VerticalAlignment)verticalAlignment).ToTextAlignmentType();
                    _.StyleFlag.VerticalAlignment = true;
                });
            }
        }

        /// <summary>
        /// Sets the horizontal alignment of a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="horizontalAlignment">The horizontal alignment.</param>
        /// <param name="styleContainer">The style container.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetHorizontalAlignment(
            this Range range,
            HorizontalAlignment? horizontalAlignment,
            StyleContainer styleContainer = null)
        {
            new { range }.Must().NotBeNull();

            if (horizontalAlignment != null)
            {
                range.SetStyle(styleContainer, _ =>
                {
                    _.Style.HorizontalAlignment = ((HorizontalAlignment)horizontalAlignment).ToTextAlignmentType();
                    _.StyleFlag.HorizontalAlignment = true;
                });
            }
        }

        /// <summary>
        /// Sets the row height for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="rowHeightInPixels">The row height, in pixels.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetRowHeightInPixels(
            this Range range,
            int? rowHeightInPixels)
        {
            new { range }.Must().NotBeNull();

            if (rowHeightInPixels != null)
            {
                foreach (var rowNumber in range.GetRowNumbers())
                {
                    range.Worksheet.Cells.SetRowHeightPixel(rowNumber - 1, (int)rowHeightInPixels);
                }
            }
        }

        /// <summary>
        /// Sets the column width for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="columnWidthInPixels">The column width, in pixels.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetColumnWidthInPixels(
            this Range range,
            int? columnWidthInPixels)
        {
            new { range }.Must().NotBeNull();

            if (columnWidthInPixels != null)
            {
                foreach (var columnNumber in range.GetColumnNumbers())
                {
                    range.Worksheet.Cells.SetColumnWidthPixel(columnNumber - 1, (int)columnWidthInPixels);
                }
            }
        }

        /// <summary>
        /// Autofits rows for a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="autofitRows">Use true to autofit rows.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetAutofitRows(
            this Range range,
            bool? autofitRows)
        {
            new { range }.Must().NotBeNull();

            if ((autofitRows != null) && (bool)autofitRows)
            {
                var columnNumbers = range.GetColumnNumbers();
                var minColumnNumber = columnNumbers.Min();
                var maxColumnNumber = columnNumbers.Max();
                foreach (var rowNumber in range.GetRowNumbers())
                {
                    var rowRangeCells = range.Worksheet.GetRange(rowNumber, rowNumber, minColumnNumber, maxColumnNumber).GetCells();
                    if (rowRangeCells.Any(_ => _.IsMerged))
                    {
                        var autoFitterOptions = new AutoFitterOptions { AutoFitMergedCells = true };
                        range.Worksheet.AutoFitRow(rowNumber - 1, minColumnNumber - 1, maxColumnNumber - 1, autoFitterOptions);
                    }
                    else
                    {
                        range.Worksheet.AutoFitRow(rowNumber - 1);
                    }
                }
            }
        }

        /// <summary>
        /// Merges or unmerges a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="mergeCells">True to merge cells, false to unmerge cells.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetMergeCells(
            this Range range,
            bool? mergeCells)
        {
            new { range }.Must().NotBeNull();

            if (mergeCells != null)
            {
                if ((bool)mergeCells)
                {
                    range.Merge();
                }
                else
                {
                    range.UnMerge();
                }
            }
        }

        /// <summary>
        /// Sets the inside border of a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="border">The border.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetInsideBorder(
            this Range range,
            Border border)
        {
            new { range }.Must().NotBeNull();

            if (border != null)
            {
                var cellRanges = range.GetCellRanges();
                foreach (var cellRange in cellRanges)
                {
                    var borderEdges = border.Edges.ToBorderType();
                    var borderStyle = border.Style.ToCellBorderType();
                    cellRange.SetOutlineBorder(borderEdges, borderStyle, border.Color);
                }
            }
        }

        /// <summary>
        /// Sets the outside border of a range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="border">The border.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetOutsideBorder(
            this Range range,
            Border border)
        {
            new { range }.Must().NotBeNull();

            if (border != null)
            {
                var borderEdges = border.Edges.ToBorderType();
                var borderStyle = border.Style.ToCellBorderType();
                range.SetOutlineBorder(borderEdges, borderStyle, border.Color);
            }
        }

        /// <summary>
        /// Sets a data validation.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="dataValidation">The validation styling to apply.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetDataValidation(
            this Range range,
            DataValidation dataValidation)
        {
            new { range }.Must().NotBeNull();

            if (dataValidation != null)
            {
                var validations = range.Worksheet.Validations;
                var cellArea = range.GetCellArea();
                var validation = validations[validations.Add(cellArea)];

                validation.Type = dataValidation.Kind.ToValidationType();
                validation.Operator = dataValidation.Operator.ToOperatorType();

                if (dataValidation.Operand1Value != null)
                {
                    validation.Value1 = dataValidation.Operand1Value;
                }

                if (dataValidation.Operand2Value != null)
                {
                    validation.Value2 = dataValidation.Operand2Value;
                }

                if (dataValidation.Operand1Formula != null)
                {
                    validation.Formula1 = dataValidation.Operand1Formula;
                }

                if (dataValidation.Operand2Formula != null)
                {
                    validation.Formula2 = dataValidation.Operand2Formula;
                }

                validation.IgnoreBlank = dataValidation.IgnoreBlank;
                validation.ShowInput = dataValidation.ShowInputMessage;
                validation.InputTitle = dataValidation.InputMessageTitle;
                validation.InputMessage = dataValidation.InputMessageBody;
                validation.ShowError = dataValidation.ShowErrorAlertAfterInvalidDataIsEntered;
                validation.AlertStyle = dataValidation.ErrorAlertStyle.ToValidationAlertType();
                validation.ErrorTitle = dataValidation.ErrorAlertTitle;
                validation.ErrorMessage = dataValidation.ErrorAlertBody;
                validation.InCellDropDown = dataValidation.ShowListDropdown;
            }
        }

        /// <summary>
        /// Sets a cell-value-based conditional formatting.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="rules">The conditional formatting rules to apply.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
        public static void SetCellValueConditionalFormatting(
            this Range range,
            IReadOnlyList<CellValueConditionalFormattingRule> rules)
        {
            new { range }.Must().NotBeNull();

            if (rules != null)
            {
                var formatConditionsIndex = range.Worksheet.ConditionalFormattings.Add();
                var formatConditions = range.Worksheet.ConditionalFormattings[formatConditionsIndex];
                var cellArea = range.GetCellArea();
                formatConditions.AddArea(cellArea);

                foreach (var rule in rules)
                {
                    var operatorType = rule.Operator.ToOperatorType();
                    var conditionIndex = formatConditions.AddCondition(FormatConditionType.CellValue, operatorType, rule.Operand1Formula, rule.Operand2Formula);
                    var formatCondition = formatConditions[conditionIndex];

                    formatCondition.StopIfTrue = rule.StopIfTrue;

                    // need a way to leverage our Set... methods above to inflate the style.
                    var backgroundColor = rule.RangeStyle?.BackgroundColor;
                    if (backgroundColor != null)
                    {
                        formatCondition.Style.BackgroundColor = (Color)backgroundColor;
                    }
                }
            }
        }

        private static void SetStyle(
            this Range range,
            StyleContainer styleContainer,
            Action<StyleContainer> configureStyleContainer)
        {
            new { range }.Must().NotBeNull();
            new { configureStyleContainer }.Must().NotBeNull();

            var applyToRange = styleContainer == null;
            if (styleContainer == null)
            {
                styleContainer = StyleContainer.BuildNewFromRange(range);
            }

            configureStyleContainer(styleContainer);

            if (applyToRange)
            {
                styleContainer.ApplyToRange(range);
            }
        }
    }
}
