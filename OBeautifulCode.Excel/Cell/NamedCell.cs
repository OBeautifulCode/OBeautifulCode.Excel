// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedCell.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Type;

    using static System.FormattableString;

    /// <summary>
    /// A cell that is referenced by a user-specified name.
    /// </summary>
    /// <remarks>
    /// Like an Excel named range, except scoped to a cell.
    /// </remarks>
    public partial class NamedCell : IModelViaCodeGen
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
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(Invariant($"'{nameof(name)}' is white space"));
            }

            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            this.Name = name;
            this.Cell = cell;
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the cell that is being named.
        /// </summary>
        public CellReference Cell { get; private set; }
    }
}
