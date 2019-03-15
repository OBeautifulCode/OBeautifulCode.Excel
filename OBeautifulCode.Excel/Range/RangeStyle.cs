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
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "This is not excessively complex; there are a lot of ways to style a range.")]
        public static bool operator ==(
            RangeStyle item1,
            RangeStyle item2)
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
                (item1.BackgroundColor == item2.BackgroundColor) &&
                (item1.FontColor == item2.FontColor) &&
                (item1.FontName == item2.FontName) &&
                (item1.FontSize == item2.FontSize) &&
                (item1.FontIsItalic == item2.FontIsItalic) &&
                (item1.FontIsBold == item2.FontIsBold) &&
                (item1.FontUnderline == item2.FontUnderline) &&
                (item1.FontRotationAngle == item2.FontRotationAngle) &&
                (item1.TextIsWrapped == item2.TextIsWrapped) &&
                (item1.IndentLevel == item2.IndentLevel) &&
                (item1.RowHeightInPixels == item2.RowHeightInPixels) &&
                (item1.ColumnWidthInPixels == item2.ColumnWidthInPixels) &&
                (item1.VerticalAlignment == item2.VerticalAlignment) &&
                (item1.HorizontalAlignment == item2.HorizontalAlignment) &&
                (item1.MergeCells == item2.MergeCells) &&
                (item1.AutofitRows == item2.AutofitRows) &&
                (item1.InsideBorder == item2.InsideBorder) &&
                (item1.OutsideBorder == item2.OutsideBorder) &&
                (item1.DataValidation == item2.DataValidation) &&
                (item1.Format == item2.Format) &&
                (item1.CustomFormatString == item2.CustomFormatString);

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="RangeStyle"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            RangeStyle item1,
            RangeStyle item2)
            => !(item1 == item2);

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
