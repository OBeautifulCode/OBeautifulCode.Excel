// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TextDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Type;

    /// <summary>
    /// Represents validation against textual data.
    /// </summary>
    /// <remarks>
    /// This class could be used to validate data that not textual because Excel
    /// is pretty loose on data type conversions.  For example, the user could enter
    /// a date and the validation could compare that date against the string "1/1/2000",
    /// which Excel will interpret as a date (perhaps this only works with a string?).
    /// </remarks>
    public partial class TextDataValidation : DataValidation, IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the first operand value.
        /// </summary>
        public string Operand1TextValue { get; set; }

        /// <summary>
        /// Gets or sets the second operand value.
        /// </summary>
        public string Operand2TextValue { get; set; }

        /// <inheritdoc />
        public override object Operand1Value => this.Operand1TextValue;

        /// <inheritdoc />
        public override object Operand2Value => this.Operand2TextValue;
    }
}
