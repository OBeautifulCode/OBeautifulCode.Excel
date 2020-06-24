// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedCell.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Assertion.Recipes;
    using OBeautifulCode.Equality.Recipes;
    using OBeautifulCode.Type;

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
            new { name }.AsArg().Must().NotBeNullNorWhiteSpace();
            new { cell }.AsArg().Must().NotBeNull();

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
