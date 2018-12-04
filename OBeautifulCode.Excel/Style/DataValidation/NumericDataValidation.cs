﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NumericDataValidation.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Linq.Expressions;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Represents validation against numeric data.
    /// </summary>
    public class NumericDataValidation : DataValidation, IEquatable<NumericDataValidation>
    {
        private static readonly Func<NumericDataValidation, NumericDataValidation> CloneFunc =
            MappingExpression.From<NumericDataValidation>.ToNew<NumericDataValidation>().Compile();

        /// <summary>
        /// Gets or sets the first operand value.
        /// </summary>
        public long? Operand1Value { get; set; }

        /// <summary>
        /// Gets or sets the second operand value.
        /// </summary>
        public long? Operand2Value { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="NumericDataValidation"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            NumericDataValidation item1,
            NumericDataValidation item2)
        {
            var result = Equals(item1, item2);
            if (result && !ReferenceEquals(item1, null))
            {
                // ReSharper disable once PossibleNullReferenceException
                result = (item1.Operand1Value == item2.Operand1Value) &&
                         (item1.Operand2Value == item2.Operand2Value);
            }

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="NumericDataValidation"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            NumericDataValidation item1,
            NumericDataValidation item2)
            => !(item1 == item2);

        /// <inheritdoc />
        public bool Equals(NumericDataValidation other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as NumericDataValidation);

        /// <inheritdoc />
        public override int GetHashCode() =>
            new HashCodeHelper(GetHashCode(this))
                .Hash(this.Operand1Value)
                .Hash(this.Operand2Value)
                .Value;

        /// <inheritdoc />
        public override DataValidation Clone()
        {
            var result = CloneFunc(this);
            return result;
        }
    }
}