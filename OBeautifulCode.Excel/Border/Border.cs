// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Border.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Drawing;

    using OBeautifulCode.Type;

    /// <summary>
    /// Defines a border.
    /// </summary>
    public partial class Border : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the edges of the border.
        /// </summary>
        public BorderEdges Edges { get; set; }

        /// <summary>
        /// Gets or sets the style of the border.
        /// </summary>
        public BorderStyle Style { get; set; }

        /// <summary>
        /// Gets or sets the color of the border.
        /// </summary>
        public Color Color { get; set; }
    }
}
