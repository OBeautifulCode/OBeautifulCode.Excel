// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Comment.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Drawing;
    using System.Linq.Expressions;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Represents a comment on a cell.
    /// </summary>
    /// <remarks>
    /// Although many of the properties of this type overlap with <see cref="RangeStyle"/>, comments are their
    /// own thing.  For example, comments have a height and width, not a row height and column width, as in <see cref="RangeStyle"/>.
    /// Also, there are fewer kinds of formatting that can be applied to a comment.
    /// We have deliberately chosen NOT to consolidate many of the properties below into a single <see cref="RangeStyle"/> property.
    /// </remarks>
    public class Comment : IEquatable<Comment>
    {
        private static readonly Func<Comment, Comment> CloneFunc = MappingExpression.From<Comment>.ToNew<Comment>().Compile();

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
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            Comment item1,
            Comment item2)
        {
            if (ReferenceEquals(item1, item2))
            {
                return true;
            }

            if (ReferenceEquals(item1, null) || ReferenceEquals(item2, null))
            {
                return false;
            }

            var result =
                (item1.Body == item2.Body) &&
                (item1.HtmlBody == item2.HtmlBody) &&
                (item1.FontName == item2.FontName) &&
                (item1.FontColor == item2.FontColor) &&
                (item1.FontSize == item2.FontSize) &&
                (item1.FontIsBold == item2.FontIsBold) &&
                (item1.HorizontalAlignment == item2.HorizontalAlignment) &&
                (item1.VerticalAlignment == item2.VerticalAlignment) &&
                (item1.AutoSize == item2.AutoSize) &&
                (item1.HeightInInches == item2.HeightInInches) &&
                (item1.WidthInInches == item2.WidthInInches) &&
                (item1.FillColor == item2.FillColor) &&
                (item1.FillTransparency == item2.FillTransparency) &&
                (item1.BorderColor == item2.BorderColor) &&
                (item1.BorderStyle == item2.BorderStyle) &&
                (item1.BorderWeightInPoints == item2.BorderWeightInPoints);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="Comment"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            Comment item1,
            Comment item2)
            => !(item1 == item2);

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

        /// <summary>
        /// Creates a clone of this object.
        /// </summary>
        /// <returns>
        /// A clone of this object.
        /// </returns>
        public Comment Clone()
        {
            var result = CloneFunc(this);

            return result;
        }
    }
}
