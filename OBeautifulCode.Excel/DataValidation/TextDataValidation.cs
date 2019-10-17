﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TextDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.Equality.Recipes;

    /// <summary>
    /// Represents validation against textual data.
    /// </summary>
    /// <remarks>
    /// This class could be used to validate data that not textual because Excel
    /// is pretty loose on data type conversions.  For example, the user could enter
    /// a date and the validation could compare that date against the string "1/1/2000",
    /// which Excel will interpret as a date (perhaps this only works with a string?).
    /// </remarks>
    public class TextDataValidation : DataValidation, IEquatable<TextDataValidation>
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

        /// <summary>
        /// Determines whether two objects of type <see cref="TextDataValidation"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1", Justification = "This is being validated via base.Equals")]
        public static bool operator ==(
            TextDataValidation left,
            TextDataValidation right)
        {
            var result = Equals(left, right);
            if (result && !ReferenceEquals(left, null))
            {
                // ReSharper disable once PossibleNullReferenceException
                result = (left.Operand1TextValue == right.Operand1TextValue) &&
                         (left.Operand2TextValue == right.Operand2TextValue);
            }

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="TextDataValidation"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            TextDataValidation left,
            TextDataValidation right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(TextDataValidation other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as TextDataValidation);

        /// <inheritdoc />
        public override int GetHashCode() =>
            new HashCodeHelper(GetHashCode(this))
                .Hash(this.Operand1TextValue)
                .Hash(this.Operand2TextValue)
                .Value;

        /// <inheritdoc />
        public override DataValidation DeepClone()
        {
            var result = new TextDataValidation
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
                Operand1TextValue = this.Operand1TextValue,
                Operand2TextValue = this.Operand2TextValue,
            };

            return result;
        }
    }
}
