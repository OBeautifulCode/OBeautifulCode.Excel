﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellExtensions.Write.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Net;

    using Aspose.Cells;
    using Aspose.Cells.Drawing;

    using static System.FormattableString;

    using Comment = OBeautifulCode.Excel.Comment;
    using Range = Aspose.Cells.Range;

    /// <summary>
    /// Extension methods on type <see cref="Cell"/>.
    /// </summary>
    public static partial class CellExtensions
    {
        /// <summary>
        /// Inserts a set of images into a cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="imageUrls">The URLs of the images to insert.</param>
        /// <param name="imageWidthScale">Optional scale to use for the image width.  Default is 100 which maintains the image's original width.  Lower is smaller; higher is larger.</param>
        /// <param name="imageHeightScale">Optional scale to use for the image height.  Default is 100 which maintains the image's original height.  Lower is smaller; higher is larger.</param>
        /// <param name="relativeOrientation">Optional specification of the orientation of images relative to each other.  Default is horizontal.  Doesn't matter if there is only one image.</param>
        /// <param name="cellSizeChanges">Optional specification of the changes to make to the size of a cell to fit the images.  Default is to expand both the row and column to fit the images.</param>
        /// <param name="rowHeightInPixels">Optional fixed height to use for all rows that the image overlaps with, when <paramref name="cellSizeChanges"/> is <see cref="ImagesCellSizeChanges.ResizeRowsToFixedHeight"/>.  Default is <see cref="Constants.DefaultRowHeightInPixels"/>.</param>
        /// <param name="columnWidthInPixels">Optional fixed width to use for all columns that the image overlaps with, when <paramref name="cellSizeChanges"/> is <see cref="ImagesCellSizeChanges.ResizeColumnsToFixedWidth"/>.  Default is <see cref="Constants.DefaultColumnWidthInPixels"/>.</param>
        /// <param name="autoLayoutProcedures">Optional specification of the automatic layout procedures to apply to the images.  Default is to auto-space and auto-align the images.</param>
        /// <returns>
        /// An object containing the results of inserting the images.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="imageUrls"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageUrls"/> is empty.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageUrls"/> contains a null or white space element.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageWidthScale"/> is less than 1 or greater than 500.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageHeightScale"/> is less than 1 or greater than 500.</exception>
        /// <exception cref="ArgumentException"><paramref name="relativeOrientation"/> is <see cref="ImagesRelativeOrientation.Unknown"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="rowHeightInPixels"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnWidthInPixels"/> is less than 1.</exception>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex.")]
        [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Justification = "This is not excessively coupled.")]
        public static InsertImagesResult InsertImages(
            this Cell cell,
            IReadOnlyList<string> imageUrls,
            int imageWidthScale = 100,
            int imageHeightScale = 100,
            ImagesRelativeOrientation relativeOrientation = ImagesRelativeOrientation.Horizontal,
            ImagesCellSizeChanges cellSizeChanges = ImagesCellSizeChanges.ExpandRowAndColumnToFitImages,
            int rowHeightInPixels = Constants.DefaultRowHeightInPixels,
            int columnWidthInPixels = Constants.DefaultColumnWidthInPixels,
            ImagesAutoLayoutProcedures autoLayoutProcedures = ImagesAutoLayoutProcedures.AutoSpaceAndAutoAlign)
        {
            if (imageUrls == null)
            {
                throw new ArgumentNullException(nameof(imageUrls));
            }

            if (!imageUrls.Any())
            {
                throw new ArgumentException(Invariant($"'{nameof(imageUrls)}' is an empty enumerable"));
            }

            if (imageUrls.Any(string.IsNullOrWhiteSpace))
            {
                throw new ArgumentException(Invariant($"{nameof(imageUrls)} contains an element that is null or white space"));
            }

            var imagesBytes = new List<byte[]>();

            using (var webClient = new WebClient())
            {
                foreach (var imageUrl in imageUrls)
                {
                    var imageBytes = webClient.DownloadData(imageUrl);

                    imagesBytes.Add(imageBytes);
                }
            }

            var result = cell.InsertImages(
                imagesBytes,
                imageWidthScale,
                imageHeightScale,
                relativeOrientation,
                cellSizeChanges,
                rowHeightInPixels,
                columnWidthInPixels,
                autoLayoutProcedures);

            return result;
        }

        /// <summary>
        /// Inserts a set of images into a cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="imagesBytes">The bytes of the images to insert.</param>
        /// <param name="imageWidthScale">Optional scale to use for the image width.  Default is 100 which maintains the image's original width.  Lower is smaller; higher is larger.</param>
        /// <param name="imageHeightScale">Optional scale to use for the image height.  Default is 100 which maintains the image's original height.  Lower is smaller; higher is larger.</param>
        /// <param name="relativeOrientation">Optional specification of the orientation of images relative to each other.  Default is horizontal.  Doesn't matter if there is only one image.</param>
        /// <param name="cellSizeChanges">Optional specification of the changes to make to the size of a cell to fit the images.  Default is to expand both the row and column to fit the images.</param>
        /// <param name="rowHeightInPixels">Optional fixed height to use for all rows that the image overlaps with, when <paramref name="cellSizeChanges"/> is <see cref="ImagesCellSizeChanges.ResizeRowsToFixedHeight"/>.  Default is <see cref="Constants.DefaultRowHeightInPixels"/>.</param>
        /// <param name="columnWidthInPixels">Optional fixed width to use for all columns that the image overlaps with, when <paramref name="cellSizeChanges"/> is <see cref="ImagesCellSizeChanges.ResizeColumnsToFixedWidth"/>.  Default is <see cref="Constants.DefaultColumnWidthInPixels"/>.</param>
        /// <param name="autoLayoutProcedures">Optional specification of the automatic layout procedures to apply to the images.  Default is to auto-space and auto-align the images.</param>
        /// <returns>
        /// An object containing the results of inserting the images.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="imagesBytes"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="imagesBytes"/> is empty.</exception>
        /// <exception cref="ArgumentException"><paramref name="imagesBytes"/> contains a null or empty element.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageWidthScale"/> is less than 1 or greater than 500.</exception>
        /// <exception cref="ArgumentException"><paramref name="imageHeightScale"/> is less than 1 or greater than 500.</exception>
        /// <exception cref="ArgumentException"><paramref name="relativeOrientation"/> is <see cref="ImagesRelativeOrientation.Unknown"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="rowHeightInPixels"/> is less than 1.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnWidthInPixels"/> is less than 1.</exception>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex.")]
        [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Justification = "This is not excessively coupled.")]
        public static InsertImagesResult InsertImages(
            this Cell cell,
            IReadOnlyList<byte[]> imagesBytes,
            int imageWidthScale = 100,
            int imageHeightScale = 100,
            ImagesRelativeOrientation relativeOrientation = ImagesRelativeOrientation.Horizontal,
            ImagesCellSizeChanges cellSizeChanges = ImagesCellSizeChanges.ExpandRowAndColumnToFitImages,
            int rowHeightInPixels = Constants.DefaultRowHeightInPixels,
            int columnWidthInPixels = Constants.DefaultColumnWidthInPixels,
            ImagesAutoLayoutProcedures autoLayoutProcedures = ImagesAutoLayoutProcedures.AutoSpaceAndAutoAlign)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            if (imagesBytes == null)
            {
                throw new ArgumentNullException(nameof(imagesBytes));
            }

            if (!imagesBytes.Any())
            {
                throw new ArgumentException(Invariant($"'{nameof(imagesBytes)}' is an empty enumerable"));
            }

            if (imagesBytes.Any(_ => _ == null))
            {
                throw new ArgumentException(Invariant($"{nameof(imagesBytes)} contains an element that is null"));
            }

            if (imagesBytes.Any(_ => _.Length == 0))
            {
                throw new ArgumentException(Invariant($"{nameof(imagesBytes)} contains an element that is empty"));
            }

            if (imageWidthScale < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(imageWidthScale)}' < '{1}'"), (Exception)null);
            }

            if (imageWidthScale > 500)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(imageWidthScale)}' > '{500}'"), (Exception)null);
            }

            if (imageHeightScale < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(imageHeightScale)}' < '{1}'"), (Exception)null);
            }

            if (imageHeightScale > 500)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(imageHeightScale)}' > '{500}'"), (Exception)null);
            }

            if (relativeOrientation == ImagesRelativeOrientation.Unknown)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(relativeOrientation)}' == '{ImagesRelativeOrientation.Unknown}'"), (Exception)null);
            }

            if (rowHeightInPixels < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(rowHeightInPixels)}' < '{1}'"), (Exception)null);
            }

            if (columnWidthInPixels < 1)
            {
                throw new ArgumentOutOfRangeException(Invariant($"'{nameof(columnWidthInPixels)}' < '{1}'"), (Exception)null);
            }

            if (relativeOrientation == ImagesRelativeOrientation.Vertical)
            {
                throw new NotImplementedException(Invariant($"This {nameof(ImagesRelativeOrientation)} is not yet implemented: {nameof(ImagesRelativeOrientation.Vertical)}"));
            }

            var worksheet = cell.Worksheet;

            var pictures = new List<Picture>();

            var pictureIndices = new List<int>();

            foreach (var imageBytes in imagesBytes)
            {
                using (var imageStream = new MemoryStream(imageBytes))
                {
                    var pictureIndex = worksheet.Pictures.Add(cell.Row, cell.Column, imageStream);

                    pictureIndices.Add(pictureIndex);

                    var picture = worksheet.Pictures[pictureIndex];

                    var imageBitmap = new Bitmap(imageStream);
                    var imageHeightInPixels = imageBitmap.Height;
                    var imageWidthInPixels = imageBitmap.Width;

                    picture.Height = imageHeightInPixels;
                    picture.Width = imageWidthInPixels;
                    picture.HeightScale = imageHeightScale;
                    picture.WidthScale = imageWidthScale;
                    pictures.Add(picture);
                }
            }

            var maxImageWidthInPixels = pictures.Max(_ => _.Width);
            var maxImageHeightInPixels = pictures.Max(_ => _.Height);
            var totalImageWidthInPixels = pictures.Sum(_ => _.Width);

            int resultEndRow, resultEndColumn;
            var resultStartRow = resultEndRow = cell.GetRowNumber();
            var resultStartColumn = resultEndColumn = cell.GetColumnNumber();

            // setting the row height or column height could move the pictures so do that before positioning the pictures
            if (cellSizeChanges.HasFlag(ImagesCellSizeChanges.ExpandRowToFitImages))
            {
                var rowHeight = cell.GetHeightInPixels();
                if (maxImageHeightInPixels > rowHeight)
                {
                    cell.GetRange().SetTotalRowHeightInPixels(maxImageHeightInPixels);
                }
            }
            else
            {
                var cursor = new CellCursor(worksheet, cell.GetRowNumber(), cell.GetColumnNumber());
                var pixelsCovered = 0;

                do
                {
                    if (cellSizeChanges.HasFlag(ImagesCellSizeChanges.ResizeRowsToFixedHeight))
                    {
                        cursor.CellRange.SetPerRowHeightInPixels(rowHeightInPixels);
                    }

                    pixelsCovered += cursor.Cell.GetHeightInPixels(includeMergedCells: false);
                    cursor.MoveDown();
                }
                while (pixelsCovered < maxImageHeightInPixels);

                resultEndRow = cursor.RowNumber - 1;
            }

            if (cellSizeChanges.HasFlag(ImagesCellSizeChanges.ExpandColumnToFitImages))
            {
                var columnWidth = cell.GetWidthInPixels();

                if (autoLayoutProcedures.HasFlag(ImagesAutoLayoutProcedures.AutoSpace))
                {
                    if (totalImageWidthInPixels > columnWidth)
                    {
                        cell.GetRange().SetTotalColumnWidthInPixels(totalImageWidthInPixels);
                    }
                }
                else
                {
                    if (maxImageWidthInPixels > columnWidth)
                    {
                        cell.GetRange().SetTotalColumnWidthInPixels(maxImageWidthInPixels);
                    }
                }
            }
            else
            {
                var cursor = new CellCursor(worksheet, cell.GetRowNumber(), cell.GetColumnNumber());
                var pixelsCovered = 0;

                do
                {
                    if (cellSizeChanges.HasFlag(ImagesCellSizeChanges.ResizeColumnsToFixedWidth))
                    {
                        cursor.CellRange.SetPerColumnWidthInPixels(columnWidthInPixels);
                    }

                    pixelsCovered += cursor.Cell.GetWidthInPixels(includeMergedCells: false);
                    cursor.MoveRight();
                }
                while (pixelsCovered < maxImageWidthInPixels);

                resultEndColumn = cursor.ColumnNumber - 1;
            }

            if (autoLayoutProcedures.HasFlag(ImagesAutoLayoutProcedures.AutoSpace))
            {
                var horizontalMarginInPixels = (cell.GetWidthInPixels() - totalImageWidthInPixels) / (pictures.Count + 1);
                var horizontalPositionInPixels = pictures[0].X;

                foreach (var picture in pictures)
                {
                    if (horizontalMarginInPixels >= 0)
                    {
                        picture.X = horizontalPositionInPixels + horizontalMarginInPixels;
                        horizontalPositionInPixels = picture.X + picture.Width;
                    }
                }
            }

            if (autoLayoutProcedures.HasFlag(ImagesAutoLayoutProcedures.AutoAlign))
            {
                foreach (var picture in pictures)
                {
                    picture.Y += (maxImageHeightInPixels - picture.Height) / 2;
                }
            }

            var range = worksheet.GetRange(resultStartRow, resultEndRow, resultStartColumn, resultEndColumn);

            var result = new InsertImagesResult(range, pictureIndices);

            return result;
        }

        /// <summary>
        /// Freezes panes at the specified cell.
        /// </summary>
        /// <remarks>
        /// The space above and/or to the left of the specified cell will be frozen, based on <paramref name="paneKinds"/>.
        /// </remarks>
        /// <param name="cell">The cell at which to freeze panes.</param>
        /// <param name="paneKinds">The kinds of panes to freeze.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cell"/> is null.</exception>
        public static void SetFreezePanes(
            this Cell cell,
            PaneKinds paneKinds)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

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
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

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
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

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
