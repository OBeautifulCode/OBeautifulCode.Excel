﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Comment.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Drawing;

    using OBeautifulCode.Type;

    /// <summary>
    /// Represents a comment on a cell.
    /// </summary>
    /// <remarks>
    /// Although many of the properties of this type overlap with <see cref="RangeStyle"/>, comments are their
    /// own thing.  For example, comments have a height and width, not a row height and column width, as in <see cref="RangeStyle"/>.
    /// Also, there are fewer kinds of formatting that can be applied to a comment.
    /// We have deliberately chosen NOT to consolidate many of the properties below into a single <see cref="RangeStyle"/> property.
    /// </remarks>
    public partial class Comment : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the body of the comment.
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets the HTML body of the comment.
        /// </summary>
        public string HtmlBody { get; set; }

        /// <summary>
        /// Gets or sets the font name.
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Gets or sets the font color.
        /// </summary>
        public Color? FontColor { get; set; }

        /// <summary>
        /// Gets or sets the size of the font.
        /// </summary>
        public int? FontSize { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the font is bold or not.
        /// </summary>
        public bool? FontIsBold { get; set; }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public VerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to auto-size the comment.
        /// </summary>
        public bool? AutoSize { get; set; }

        /// <summary>
        /// Gets or sets the height of the comment, in inches.
        /// </summary>
        public decimal? HeightInInches { get; set; }

        /// <summary>
        /// Gets or sets the width of the comment, in inches.
        /// </summary>
        public decimal? WidthInInches { get; set; }

        /// <summary>
        /// Gets or sets the fill color.
        /// </summary>
        public Color? FillColor { get; set; }

        /// <summary>
        /// Gets or sets the fill transparency.
        /// </summary>
        public decimal? FillTransparency { get; set; }

        /// <summary>
        /// Gets or sets the color of the border.
        /// </summary>
        public Color? BorderColor { get; set; }

        /// <summary>
        /// Gets or sets the border style.
        /// </summary>
        public CommentBorderStyle? BorderStyle { get; set; }

        /// <summary>
        /// Gets or sets the weight of the border, in points.
        /// </summary>
        public decimal? BorderWeightInPoints { get; set; }
    }
}
