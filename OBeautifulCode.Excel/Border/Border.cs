// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Border.cs" company="OBeautifulCode">
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
    /// Defines a border.
    /// </summary>
    public class Border : IEquatable<Border>, IDeepCloneable<Border>
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
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            Border left,
            Border right)
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
                (left.Edges == right.Edges) &&
                (left.Style == right.Style) &&
                (left.Color == right.Color);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="Border"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            Border left,
            Border right)
            => !(left == right);

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

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
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
