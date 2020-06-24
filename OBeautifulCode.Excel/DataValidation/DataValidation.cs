// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Type;

    /// <summary>
    /// Represents validation applied to data entered by a user.
    /// </summary>
    public abstract partial class DataValidation : IModelViaCodeGen
    {
        /// <summary>
        /// Gets or sets the kind of validation to perform.
        /// </summary>
        public DataValidationKind Kind { get; set; }

        /// <summary>
        /// Gets or sets the operator to apply to the data.
        /// </summary>
        public DataValidationOperator Operator { get; set; }

        /// <summary>
        /// Gets or sets the first operand formula.
        /// </summary>
        public string Operand1Formula { get; set; }

        /// <summary>
        /// Gets or sets the second operand formula.
        /// </summary>
        public string Operand2Formula { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether blanks should be ignored.
        /// </summary>
        public bool IgnoreBlank { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to show an input message.
        /// </summary>
        public bool ShowInputMessage { get; set; }

        /// <summary>
        /// Gets or sets the input message title.
        /// </summary>
        public string InputMessageTitle { get; set; }

        /// <summary>
        /// Gets or sets the input message body.
        /// </summary>
        public string InputMessageBody { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether an error alert should be shown after invalid data is entered.
        /// </summary>
        public bool ShowErrorAlertAfterInvalidDataIsEntered { get; set; }

        /// <summary>
        /// Gets or sets the style of the error alert.
        /// </summary>
        public DataValidationErrorAlertStyle ErrorAlertStyle { get; set; }

        /// <summary>
        /// Gets or sets the title of the error alert.
        /// </summary>
        public string ErrorAlertTitle { get; set; }

        /// <summary>
        /// Gets or sets the body of the error alert.
        /// </summary>
        public string ErrorAlertBody { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a list dropdown should be shown.
        /// </summary>
        public bool ShowListDropdown { get; set; }

        /// <summary>
        /// Gets the first operand value.
        /// </summary>
        public abstract object Operand1Value { get; }

        /// <summary>
        /// Gets the second operand value.
        /// </summary>
        public abstract object Operand2Value { get; }
    }
}
