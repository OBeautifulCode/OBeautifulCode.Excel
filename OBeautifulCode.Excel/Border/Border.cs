// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Border.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Drawing;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Defines a border.
    /// </summary>
    public class Border : IEquatable<Border>
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

        /// <summary>
        /// Determines whether two objects of type <see cref="Border"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            Border item1,
            Border item2)
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
                (item1.Edges == item2.Edges) &&
                (item1.Style == item2.Style) &&
                (item1.Color == item2.Color);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="Border"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            Border item1,
            Border item2)
            => !(item1 == item2);

        /// <inheritdoc />
        public bool Equals(Border other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as Border);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.Edges)
                .Hash(this.Style)
                .Hash(this.Color)
                .Value;

        /// <summary>
        /// Creates a deep clone of this object.
        /// </summary>
        /// <returns>
        /// A deep clone of this object.
        /// </returns>
        public Border DeepClone()
        {
            var result = new Border
            {
                Edges = this.Edges,
                Style = this.Style,
                Color = this.Color,
            };

            return result;
        }
    }
}
