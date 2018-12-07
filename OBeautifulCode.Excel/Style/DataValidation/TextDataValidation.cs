// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TextDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq.Expressions;

    using OBeautifulCode.Math.Recipes;

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
        private static readonly Func<TextDataValidation, TextDataValidation> CloneFunc =
            MappingExpression.From<TextDataValidation>.ToNew<TextDataValidation>().Compile();

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
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1", Justification = "This is being validated via base.Equals")]
        public static bool operator ==(
            TextDataValidation item1,
            TextDataValidation item2)
        {
            var result = Equals(item1, item2);
            if (result && !ReferenceEquals(item1, null))
            {
                // ReSharper disable once PossibleNullReferenceException
                result = (item1.Operand1TextValue == item2.Operand1TextValue) &&
                         (item1.Operand2TextValue == item2.Operand2TextValue);
            }

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="TextDataValidation"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            TextDataValidation item1,
            TextDataValidation item2)
            => !(item1 == item2);

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
        public override DataValidation Clone()
        {
            var result = CloneFunc(this);
            return result;
        }
    }
}
