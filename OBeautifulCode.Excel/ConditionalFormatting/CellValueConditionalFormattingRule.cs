// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellValueConditionalFormattingRule.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.Equality.Recipes;
    using OBeautifulCode.Type;

    /// <summary>
    /// Specifies a conditional formatting rule based on the value of a cell.
    /// </summary>
    public class CellValueConditionalFormattingRule : IEquatable<CellValueConditionalFormattingRule>, IDeepCloneable<CellValueConditionalFormattingRule>
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
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            CellValueConditionalFormattingRule left,
            CellValueConditionalFormattingRule right)
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
                (left.Operator == right.Operator) &&
                (left.Operand1Formula == right.Operand1Formula) &&
                (left.Operand2Formula == right.Operand2Formula) &&
                (left.StopIfTrue == right.StopIfTrue) &&
                (left.RangeStyle == right.RangeStyle);
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellValueConditionalFormattingRule"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            CellValueConditionalFormattingRule left,
            CellValueConditionalFormattingRule right)
            => !(left == right);

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

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public CellValueConditionalFormattingRule DeepClone()
        {
            var result = new CellValueConditionalFormattingRule
            {
                Operator = this.Operator,
                Operand1Formula = this.Operand1Formula,
                Operand2Formula = this.Operand2Formula,
                StopIfTrue = this.StopIfTrue,
                RangeStyle = this.RangeStyle?.DeepClone(),
            };

            return result;
        }
    }
}
