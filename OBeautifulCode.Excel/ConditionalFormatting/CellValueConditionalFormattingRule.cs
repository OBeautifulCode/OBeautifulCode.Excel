// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellValueConditionalFormattingRule.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// Specifies a conditional formatting rule based on the value of a cell.
    /// </summary>
    public class CellValueConditionalFormattingRule : IEquatable<CellValueConditionalFormattingRule>
    {
        /// <summary>
        /// Gets or sets the operator to use.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords", MessageId = "Operator", Justification = "This is the best name for this property.")]
        public ConditionalFormattingOperator Operator { get; set; }

        /// <summary>
        /// Gets or sets the first operand formula.
        /// </summary>
        public string Operand1Formula { get; set; }

        /// <summary>
        /// Gets or sets the second operand formula.
        /// </summary>
        public string Operand2Formula { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to stop
        /// evaluating conditions if the condition is true.
        /// </summary>
        public bool StopIfTrue { get; set; }

        /// <summary>
        /// Gets or sets the style to apply when the condition is true.
        /// </summary>
        public RangeStyle RangeStyle { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellValueConditionalFormattingRule"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            CellValueConditionalFormattingRule item1,
            CellValueConditionalFormattingRule item2)
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
                (item1.Operator == item2.Operator) &&
                (item1.Operand1Formula == item2.Operand1Formula) &&
                (item1.Operand2Formula == item2.Operand2Formula) &&
                (item1.StopIfTrue == item2.StopIfTrue) &&
                (item1.RangeStyle == item2.RangeStyle);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellValueConditionalFormattingRule"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            CellValueConditionalFormattingRule item1,
            CellValueConditionalFormattingRule item2)
            => !(item1 == item2);

        /// <inheritdoc />
        public bool Equals(CellValueConditionalFormattingRule other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as CellValueConditionalFormattingRule);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.Operator)
                .Hash(this.Operand1Formula)
                .Hash(this.Operand2Formula)
                .Hash(this.StopIfTrue)
                .Hash(this.RangeStyle)
                .Value;

        /// <summary>
        /// Creates a clone of this object.
        /// </summary>
        /// <returns>
        /// A clone of this object.
        /// </returns>
        public CellValueConditionalFormattingRule Clone()
        {
            var result = new CellValueConditionalFormattingRule
            {
                Operator = this.Operator,
                Operand1Formula = this.Operand1Formula,
                Operand2Formula = this.Operand2Formula,
                StopIfTrue = this.StopIfTrue,
                RangeStyle = this.RangeStyle.Clone(),
            };

            return result;
        }
    }
}
