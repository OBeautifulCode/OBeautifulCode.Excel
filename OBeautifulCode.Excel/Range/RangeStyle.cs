// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeStyle.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// The style to apply to a range.
    /// </summary>
    public class RangeStyle : IEquatable<RangeStyle>
    {
        /// <summary>
        /// Gets or sets the background color.
        /// </summary>
        public Color? BackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the font color.
        /// </summary>
        public Color? FontColor { get; set; }

        /// <summary>
        /// Gets or sets the name of the font.
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Gets or sets the size of the font.
        /// </summary>
        public int? FontSize { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the font should be italics or not.
        /// </summary>
        public bool? FontIsItalic { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the font should be bold or not.
        /// </summary>
        public bool? FontIsBold { get; set; }

        /// <summary>
        /// Gets or sets the kind of underline.
        /// </summary>
        public UnderlineKind? FontUnderline { get; set; }

        /// <summary>
        /// Gets or sets the angle to rotate the font.
        /// </summary>
        public int? FontRotationAngle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether text is wrapped or not.
        /// </summary>
        public bool? TextIsWrapped { get; set; }

        /// <summary>
        /// Gets or sets the indent level.
        /// </summary>
        public int? IndentLevel { get; set; }

        /// <summary>
        /// Gets or sets the row height, in pixels. Set to 0 to hide row.
        /// </summary>
        public int? RowHeightInPixels { get; set; }

        /// <summary>
        /// Gets or sets the column width, in pixels.  Set to 0 to hide column.
        /// </summary>
        public int? ColumnWidthInPixels { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public VerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether cells should be merged.
        /// </summary>
        public bool? MergeCells { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to auto-fit rows.
        /// </summary>
        // ReSharper disable once IdentifierTypo
        public bool? AutofitRows { get; set; }

        /// <summary>
        /// Gets or sets the inside border.
        /// </summary>
        public Border InsideBorder { get; set; }

        /// <summary>
        /// Gets or sets the outside border.
        /// </summary>
        public Border OutsideBorder { get; set; }

        /// <summary>
        /// Gets or sets the data validation.
        /// </summary>
        public DataValidation DataValidation { get; set; }

        /// <summary>
        /// Gets or sets the format.
        /// </summary>
        public Format? Format { get; set; }

        /// <summary>
        /// Gets or sets the custom format string.
        /// </summary>
        public string CustomFormatString { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="RangeStyle"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex; there are a lot of ways to style a range.")]
        public static bool operator ==(
            RangeStyle left,
            RangeStyle right)
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
                (left.BackgroundColor == right.BackgroundColor) &&
                (left.FontColor == right.FontColor) &&
                (left.FontName == right.FontName) &&
                (left.FontSize == right.FontSize) &&
                (left.FontIsItalic == right.FontIsItalic) &&
                (left.FontIsBold == right.FontIsBold) &&
                (left.FontUnderline == right.FontUnderline) &&
                (left.FontRotationAngle == right.FontRotationAngle) &&
                (left.TextIsWrapped == right.TextIsWrapped) &&
                (left.IndentLevel == right.IndentLevel) &&
                (left.RowHeightInPixels == right.RowHeightInPixels) &&
                (left.ColumnWidthInPixels == right.ColumnWidthInPixels) &&
                (left.VerticalAlignment == right.VerticalAlignment) &&
                (left.HorizontalAlignment == right.HorizontalAlignment) &&
                (left.MergeCells == right.MergeCells) &&
                (left.AutofitRows == right.AutofitRows) &&
                (left.InsideBorder == right.InsideBorder) &&
                (left.OutsideBorder == right.OutsideBorder) &&
                (left.DataValidation == right.DataValidation) &&
                (left.Format == right.Format) &&
                (left.CustomFormatString == right.CustomFormatString);

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="RangeStyle"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            RangeStyle left,
            RangeStyle right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(RangeStyle other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as RangeStyle);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.BackgroundColor)
                .Hash(this.FontColor)
                .Hash(this.FontName)
                .Hash(this.FontSize)
                .Hash(this.FontIsItalic)
                .Hash(this.FontIsBold)
                .Hash(this.FontUnderline)
                .Hash(this.FontRotationAngle)
                .Hash(this.TextIsWrapped)
                .Hash(this.IndentLevel)
                .Hash(this.RowHeightInPixels)
                .Hash(this.ColumnWidthInPixels)
                .Hash(this.VerticalAlignment)
                .Hash(this.HorizontalAlignment)
                .Hash(this.MergeCells)
                .Hash(this.AutofitRows)
                .Hash(this.InsideBorder)
                .Hash(this.OutsideBorder)
                .Hash(this.DataValidation)
                .Hash(this.Format)
                .Hash(this.CustomFormatString)
                .Value;

        /// <summary>
        /// Creates a deep clone of this object.
        /// </summary>
        /// <returns>
        /// A deep clone of this object.
        /// </returns>
        public RangeStyle DeepClone()
        {
            var result = new RangeStyle
            {
                BackgroundColor = this.BackgroundColor,
                FontColor = this.FontColor,
                FontName = this.FontName,
                FontSize = this.FontSize,
                FontIsItalic = this.FontIsItalic,
                FontIsBold = this.FontIsBold,
                FontUnderline = this.FontUnderline,
                FontRotationAngle = this.FontRotationAngle,
                TextIsWrapped = this.TextIsWrapped,
                IndentLevel = this.IndentLevel,
                RowHeightInPixels = this.RowHeightInPixels,
                ColumnWidthInPixels = this.ColumnWidthInPixels,
                VerticalAlignment = this.VerticalAlignment,
                HorizontalAlignment = this.HorizontalAlignment,
                MergeCells = this.MergeCells,
                AutofitRows = this.AutofitRows,
                InsideBorder = this.InsideBorder?.DeepClone(),
                OutsideBorder = this.OutsideBorder?.DeepClone(),
                DataValidation = this.DataValidation?.DeepClone(),
                Format = this.Format,
                CustomFormatString = this.CustomFormatString,
            };

            return result;
        }
    }
}
