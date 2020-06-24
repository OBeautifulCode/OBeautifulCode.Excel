// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NumericDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Type;

    /// <summary>
    /// Represents validation against numeric data.
    /// </summary>
    public partial class NumericDataValidation : DataValidation, IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the first operand value.
        /// </summary>
        public long? Operand1NumericValue { get; set; }

        /// <summary>
        /// Gets or sets the second operand value.
        /// </summary>
        public long? Operand2NumericValue { get; set; }

        /// <inheritdoc />
        public override object Operand1Value => this.Operand1NumericValue;

        /// <inheritdoc />
        public override object Operand2Value => this.Operand2NumericValue;
    }
}
