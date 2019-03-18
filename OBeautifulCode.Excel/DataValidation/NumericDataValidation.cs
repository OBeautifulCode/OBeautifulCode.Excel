// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NumericDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Represents validation against numeric data.
    /// </summary>
    public class NumericDataValidation : DataValidation, IEquatable<NumericDataValidation>
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

        /// <summary>
        /// Determines whether two objects of type <see cref="NumericDataValidation"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1", Justification = "This is being validated via base.Equals")]
        public static bool operator ==(
            NumericDataValidation left,
            NumericDataValidation right)
        {
            var result = Equals(left, right);
            if (result && !ReferenceEquals(left, null))
            {
                // ReSharper disable once PossibleNullReferenceException
                result = (left.Operand1NumericValue == right.Operand1NumericValue) &&
                         (left.Operand2NumericValue == right.Operand2NumericValue);
            }

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="NumericDataValidation"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            NumericDataValidation left,
            NumericDataValidation right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(NumericDataValidation other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as NumericDataValidation);

        /// <inheritdoc />
        public override int GetHashCode() =>
            new HashCodeHelper(GetHashCode(this))
                .Hash(this.Operand1NumericValue)
                .Hash(this.Operand2NumericValue)
                .Value;

        /// <inheritdoc />
        public override DataValidation DeepClone()
        {
            var result = new NumericDataValidation
            {
                Kind = this.Kind,
                Operator = this.Operator,
                Operand1Formula = this.Operand1Formula,
                Operand2Formula = this.Operand2Formula,
                IgnoreBlank = this.IgnoreBlank,
                ShowInputMessage = this.ShowInputMessage,
                InputMessageTitle = this.InputMessageTitle,
                InputMessageBody = this.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = this.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = this.ErrorAlertStyle,
                ErrorAlertTitle = this.ErrorAlertTitle,
                ErrorAlertBody = this.ErrorAlertBody,
                ShowListDropdown = this.ShowListDropdown,
                Operand1NumericValue = this.Operand1NumericValue,
                Operand2NumericValue = this.Operand2NumericValue,
            };

            return result;
        }
    }
}
