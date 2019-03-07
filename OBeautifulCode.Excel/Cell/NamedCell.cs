// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedCell.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Math.Recipes;
    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// A cell that is referenced by a user-specified name.
    /// </summary>
    /// <remarks>
    /// Like an Excel named range, except scoped to a cell.
    /// </remarks>
    public class NamedCell : IEquatable<NamedCell>
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
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            NamedCell item1,
            NamedCell item2)
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
                (item1.Name == item2.Name) &&
                (item1.Cell == item2.Cell);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="NamedCell"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            NamedCell item1,
            NamedCell item2)
            => !(item1 == item2);

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
    }
}
