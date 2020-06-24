// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RangeStyle.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Drawing;

    using OBeautifulCode.Type;

    /// <summary>
    /// The style to apply to a range.
    /// </summary>
    public partial class RangeStyle : IModelViaCodeGen
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
    }
}
