// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Comment.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Drawing;

    using OBeautifulCode.Equality.Recipes;
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
    public class Comment : IEquatable<Comment>, IDeepCloneable<Comment>
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

        /// <summary>
        /// Determines whether two objects of type <see cref="Comment"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            Comment left,
            Comment right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
            {
                return false;
            }

            var result =
                (left.Body == right.Body) &&
                (left.HtmlBody == right.HtmlBody) &&
                (left.FontName == right.FontName) &&
                (left.FontColor == right.FontColor) &&
                (left.FontSize == right.FontSize) &&
                (left.FontIsBold == right.FontIsBold) &&
                (left.HorizontalAlignment == right.HorizontalAlignment) &&
                (left.VerticalAlignment == right.VerticalAlignment) &&
                (left.AutoSize == right.AutoSize) &&
                (left.HeightInInches == right.HeightInInches) &&
                (left.WidthInInches == right.WidthInInches) &&
                (left.FillColor == right.FillColor) &&
                (left.FillTransparency == right.FillTransparency) &&
                (left.BorderColor == right.BorderColor) &&
                (left.BorderStyle == right.BorderStyle) &&
                (left.BorderWeightInPoints == right.BorderWeightInPoints);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="Comment"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            Comment left,
            Comment right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(Comment other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as Comment);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.Body)
                .Hash(this.HtmlBody)
                .Hash(this.FontName)
                .Hash(this.FontColor)
                .Hash(this.FontSize)
                .Hash(this.FontIsBold)
                .Hash(this.HorizontalAlignment)
                .Hash(this.VerticalAlignment)
                .Hash(this.AutoSize)
                .Hash(this.HeightInInches)
                .Hash(this.WidthInInches)
                .Hash(this.FillColor)
                .Hash(this.FillTransparency)
                .Hash(this.BorderColor)
                .Hash(this.BorderStyle)
                .Hash(this.BorderWeightInPoints)
                .Value;

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public Comment DeepClone()
        {
            var result = new Comment
            {
                Body = this.Body,
                HtmlBody = this.HtmlBody,
                FontName = this.FontName,
                FontColor = this.FontColor,
                FontSize = this.FontSize,
                FontIsBold = this.FontIsBold,
                HorizontalAlignment = this.HorizontalAlignment,
                VerticalAlignment = this.VerticalAlignment,
                AutoSize = this.AutoSize,
                HeightInInches = this.HeightInInches,
                WidthInInches = this.WidthInInches,
                FillColor = this.FillColor,
                FillTransparency = this.FillTransparency,
                BorderColor = this.BorderColor,
                BorderStyle = this.BorderStyle,
                BorderWeightInPoints = this.BorderWeightInPoints,
            };

            return result;
        }
    }
}
