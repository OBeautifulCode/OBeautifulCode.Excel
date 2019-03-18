// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.ComponentModel;

    using OBeautifulCode.Math.Recipes;
    using OBeautifulCode.Type;
    using OBeautifulCode.Validation.Recipes;

    /// <summary>
    /// Represents validation applied to data entered by a user.
    /// </summary>
    [Bindable(true)]
    public abstract class DataValidation : IEquatable<DataValidation>, IDeepCloneable<DataValidation>
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

        /// <summary>
        /// Determines whether two objects of type <see cref="DataValidation" /> are equal.
        /// </summary>
        /// <param name="left">The first item to compare.</param>
        /// <param name="right">The second item to compare.</param>
        /// <returns>true if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            DataValidation left,
            DataValidation right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
            {
                return false;
            }

            var result = left.Equals((object)right);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="DataValidation" /> are not equal.
        /// </summary>
        /// <param name="left">The first item to compare.</param>
        /// <param name="right">The second item to compare.</param>
        /// <returns>true if the two item are not equal; false otherwise.</returns>
        public static bool operator !=(
            DataValidation left,
            DataValidation right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(
            DataValidation other)
            => this == other;

        /// <inheritdoc />
        public abstract override bool Equals(
            object obj);

        /// <inheritdoc />
        public abstract override int GetHashCode();

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public abstract DataValidation DeepClone();

        /// <summary>
        /// Determines whether two objects of type <see cref="DataValidation"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        protected static bool Equals(
            DataValidation left,
            DataValidation right)
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
                (left.Kind == right.Kind) &&
                (left.Operator == right.Operator) &&
                (left.Operand1Formula == right.Operand1Formula) &&
                (left.Operand2Formula == right.Operand2Formula) &&
                (left.IgnoreBlank == right.IgnoreBlank) &&
                (left.ShowInputMessage == right.ShowInputMessage) &&
                (left.InputMessageTitle == right.InputMessageTitle) &&
                (left.InputMessageBody == right.InputMessageBody) &&
                (left.ShowErrorAlertAfterInvalidDataIsEntered == right.ShowErrorAlertAfterInvalidDataIsEntered) &&
                (left.ErrorAlertStyle == right.ErrorAlertStyle) &&
                (left.ErrorAlertTitle == right.ErrorAlertTitle) &&
                (left.ErrorAlertBody == right.ErrorAlertBody) &&
                (left.ShowListDropdown == right.ShowListDropdown);

            return result;
        }

        /// <summary>
        /// Gets the hash code for the specified item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns>
        /// A hash code for the specified item.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="item"/> is null.</exception>
        protected static int GetHashCode(
            DataValidation item)
        {
            new { item }.Must().NotBeNull();

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
