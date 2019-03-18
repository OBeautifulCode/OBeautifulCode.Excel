// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedCell.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Math.Recipes;
    using OBeautifulCode.Type;
    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// A cell that is referenced by a user-specified name.
    /// </summary>
    /// <remarks>
    /// Like an Excel named range, except scoped to a cell.
    /// </remarks>
    public class NamedCell : IEquatable<NamedCell>, IDeepCloneable<NamedCell>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NamedCell"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="cell">The cell that is being named.</param>
        public NamedCell(
            string name,
            CellReference cell)
        {
            new { name }.Must().NotBeNullNorWhiteSpace();
            new { cell }.Must().NotBeNull();

            this.Name = name;
            this.Cell = cell;
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
        public string Name { get; private set; }

        /// <summary>
        /// Gets the cell that is being named.
        /// </summary>
        // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
        public CellReference Cell { get; private set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="NamedCell"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            NamedCell left,
            NamedCell right)
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
                (left.Name == right.Name) &&
                (left.Cell == right.Cell);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="NamedCell"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            NamedCell left,
            NamedCell right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(NamedCell other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as NamedCell);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.Name)
                .Hash(this.Cell)
                .Value;

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public NamedCell DeepClone()
        {
            var result = new NamedCell(this.Name, this.Cell.DeepClone());

            return result;
        }
    }
}
