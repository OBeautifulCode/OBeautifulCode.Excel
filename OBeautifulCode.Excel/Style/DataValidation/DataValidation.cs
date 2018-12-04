﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Represents validation applied to data entered by a user.
    /// </summary>
    public abstract class DataValidation
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
        /// Creates a clone of this object.
        /// </summary>
        /// <returns>
        /// A clone of this object.
        /// </returns>
        public abstract DataValidation Clone();

        /// <summary>
        /// Determines whether two objects of type <see cref="DataValidation"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        protected static bool Equals(
            DataValidation item1,
            DataValidation item2)
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
                (item1.Kind == item2.Kind) &&
                (item1.Operator == item2.Operator) &&
                (item1.Operand1Formula == item2.Operand1Formula) &&
                (item1.Operand2Formula == item2.Operand2Formula) &&
                (item1.IgnoreBlank == item2.IgnoreBlank) &&
                (item1.ShowInputMessage == item2.ShowInputMessage) &&
                (item1.InputMessageTitle == item2.InputMessageTitle) &&
                (item1.InputMessageBody == item2.InputMessageBody) &&
                (item1.ShowErrorAlertAfterInvalidDataIsEntered == item2.ShowErrorAlertAfterInvalidDataIsEntered) &&
                (item1.ErrorAlertStyle == item2.ErrorAlertStyle) &&
                (item1.ErrorAlertTitle == item2.ErrorAlertTitle) &&
                (item1.ErrorAlertBody == item2.ErrorAlertBody) &&
                (item1.ShowListDropdown == item2.ShowListDropdown);

            return result;
        }

        /// <summary>
        /// Gets the hash code for the specified item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns>
        /// A hash code for the specified item.
        /// </returns>
        protected static int GetHashCode(
            DataValidation item)
        {
            var result = HashCodeHelper.Initialize()
                .Hash(item.Kind)
                .Hash(item.Operator)
                .Hash(item.Operand1Formula)
                .Hash(item.Operand2Formula)
                .Hash(item.IgnoreBlank)
                .Hash(item.ShowInputMessage)
                .Hash(item.InputMessageTitle)
                .Hash(item.InputMessageBody)
                .Hash(item.ShowErrorAlertAfterInvalidDataIsEntered)
                .Hash(item.ErrorAlertStyle)
                .Hash(item.ErrorAlertTitle)
                .Hash(item.ErrorAlertBody)
                .Hash(item.ShowListDropdown)
                .Value;

            return result;
        }
    }
}