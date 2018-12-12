// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensions.Write.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Drawing;

    using Aspose.Cells;

    using OBeautifulCode.Validation.Recipes;

    using Comment = OBeautifulCode.Excel.Comment;

    /// <summary>
    /// Extension methods on type <see cref="Cell"/>.
    /// </summary>
    public static partial class CellExtensions
    {
        /// <summary>
        /// Freezes panes.
        /// </summary>
        /// <param name="cell">The cell at which to freeze panes.</param>
        /// <param name="paneKinds">The kinds of panes to freeze.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static void SetFreezePanes(
            this Cell cell,
            PaneKinds paneKinds)
        {
            new { cell }.Must().NotBeNull();

            if (paneKinds == PaneKinds.None)
            {
                cell.Worksheet.UnFreezePanes();
            }
            else
            {
                var frozenRows = paneKinds.HasFlag(PaneKinds.Row) ? cell.Row : 0;
                var frozenColumns = paneKinds.HasFlag(PaneKinds.Column) ? cell.Column : 0;
                cell.Worksheet.FreezePanes(cell.Name, frozenRows, frozenColumns);
            }
        }

        /// <summary>
        /// Unlocks a cell.
        /// </summary>
        /// <param name="cell">The cell to unlock.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static void SetUnlocked(
            this Cell cell)
        {
            new { cell }.Must().NotBeNull();

            var styleContainer = StyleContainer.BuildUsingExistingCellStyle(cell);

            styleContainer.Style.IsLocked = false;
            styleContainer.StyleFlag.Locked = true;

            styleContainer.ApplyToCell(cell);
        }

        /// <summary>
        /// Adds a comment to a cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="comment">The comment.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static void SetComment(
            this Cell cell,
            Comment comment)
        {
            new { cell }.Must().NotBeNull();

            if (comment != null)
            {
                var commentIndex = cell.Worksheet.Comments.Add(cell.Name);
                var excelComment = cell.Worksheet.Comments[commentIndex];

                // font
                if (comment.FontName != null)
                {
                    excelComment.Font.Name = comment.FontName;
                }

                if (comment.FontSize != null)
                {
                    excelComment.Font.Size = (int)comment.FontSize;
                }

                if (comment.FontColor != null)
                {
                    excelComment.Font.Color = (Color)comment.FontColor;
                }

                if (comment.FontIsBold != null)
                {
                    excelComment.Font.IsBold = (bool)comment.FontIsBold;
                }

                if (comment.HorizontalAlignment != null)
                {
                    excelComment.TextHorizontalAlignment = ((HorizontalAlignment)comment.HorizontalAlignment).ToTextAlignmentType();
                }

                if (comment.VerticalAlignment != null)
                {
                    excelComment.TextVerticalAlignment = ((VerticalAlignment)comment.VerticalAlignment).ToTextAlignmentType();
                }

                // excelComment.Font.IsBold;
                // excelComment.Font.IsItalic;
                // excelComment.Font.Underline;
                if (comment.AutoSize != null)
                {
                    excelComment.AutoSize = (bool)comment.AutoSize;
                }

                if (comment.WidthInInches != null)
                {
                    excelComment.WidthInch = (double)comment.WidthInInches;
                }

                if (comment.HeightInInches != null)
                {
                    excelComment.HeightInch = (double)comment.HeightInInches;
                }

                if (comment.FillTransparency != null)
                {
                    excelComment.CommentShape.Fill.Transparency = (double)comment.FillTransparency;
                }

                if (comment.FillColor != null)
                {
                    excelComment.CommentShape.Fill.SolidFill.Color = (Color)comment.FillColor;
                }

                if (comment.BorderColor != null)
                {
                    excelComment.CommentShape.Line.SolidFill.Color = (Color)comment.BorderColor;
                }

                if (comment.BorderWeightInPoints != null)
                {
                    excelComment.CommentShape.Line.Weight = (int)comment.BorderWeightInPoints;
                }

                if (comment.BorderWeightInPoints != null)
                {
                    excelComment.CommentShape.Line.Weight = (int)comment.BorderWeightInPoints;
                }

                if (comment.BorderStyle != null)
                {
                    excelComment.CommentShape.Line.CompoundType = ((CommentBorderStyle)comment.BorderStyle).ToMsoLineStyle();
                }

                // excelComment.CommentShape.Line.DashStyle;
                // excelComment.CommentShape.HasLine;
                if (comment.Body != null)
                {
                    excelComment.Note = comment.Body;
                }

                if (comment.HtmlBody != null)
                {
                    excelComment.HtmlNote = comment.HtmlBody;
                }
            }
        }
    }
}
