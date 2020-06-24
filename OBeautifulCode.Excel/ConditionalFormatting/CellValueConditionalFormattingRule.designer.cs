﻿// --------------------------------------------------------------------------------------------------------------------
// <auto-generated>
//   Generated using OBeautifulCode.CodeGen.ModelObject (1.0.87.0)
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using global::System;
    using global::System.CodeDom.Compiler;
    using global::System.Collections.Concurrent;
    using global::System.Collections.Generic;
    using global::System.Collections.ObjectModel;
    using global::System.Diagnostics.CodeAnalysis;
    using global::System.Globalization;
    using global::System.Linq;

    using global::OBeautifulCode.Equality.Recipes;
    using global::OBeautifulCode.Type;
    using global::OBeautifulCode.Type.Recipes;

    using static global::System.FormattableString;

    [Serializable]
    public partial class CellValueConditionalFormattingRule : IModel<CellValueConditionalFormattingRule>
    {
        /// <summary>
        /// Determines whether two objects of type <see cref="CellValueConditionalFormattingRule"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are equal; otherwise false.</returns>
        public static bool operator ==(CellValueConditionalFormattingRule left, CellValueConditionalFormattingRule right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
            {
                return false;
            }

            var result = left.Equals(right);

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="CellValueConditionalFormattingRule"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are not equal; otherwise false.</returns>
        public static bool operator !=(CellValueConditionalFormattingRule left, CellValueConditionalFormattingRule right) => !(left == right);

        /// <inheritdoc />
        public bool Equals(CellValueConditionalFormattingRule other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (ReferenceEquals(other, null))
            {
                return false;
            }

            var result = this.Operator.IsEqualTo(other.Operator)
                      && this.Operand1Formula.IsEqualTo(other.Operand1Formula, StringComparer.Ordinal)
                      && this.Operand2Formula.IsEqualTo(other.Operand2Formula, StringComparer.Ordinal)
                      && this.StopIfTrue.IsEqualTo(other.StopIfTrue)
                      && this.RangeStyle.IsEqualTo(other.RangeStyle);

            return result;
        }

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as CellValueConditionalFormattingRule);

        /// <inheritdoc />
        public override int GetHashCode() => HashCodeHelper.Initialize()
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
                                 Operator        = this.Operator,
                                 Operand1Formula = this.Operand1Formula?.Clone().ToString(),
                                 Operand2Formula = this.Operand2Formula?.Clone().ToString(),
                                 StopIfTrue      = this.StopIfTrue,
                                 RangeStyle      = this.RangeStyle?.DeepClone(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="Operator" />.
        /// </summary>
        /// <param name="operator">The new <see cref="Operator" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="CellValueConditionalFormattingRule" /> using the specified <paramref name="operator" /> for <see cref="Operator" /> and a deep clone of every other property.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002: DoNotExposeGenericLists")]
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords")]
        [SuppressMessage("Microsoft.Naming", "CA1719:ParameterNamesShouldNotMatchMemberNames")]
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames")]
        [SuppressMessage("Microsoft.Naming", "CA1722:IdentifiersShouldNotHaveIncorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration")]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly")]
        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public CellValueConditionalFormattingRule DeepCloneWithOperator(ConditionalFormattingOperator @operator)
        {
            var result = new CellValueConditionalFormattingRule
                             {
                                 Operator        = @operator,
                                 Operand1Formula = this.Operand1Formula?.Clone().ToString(),
                                 Operand2Formula = this.Operand2Formula?.Clone().ToString(),
                                 StopIfTrue      = this.StopIfTrue,
                                 RangeStyle      = this.RangeStyle?.DeepClone(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="Operand1Formula" />.
        /// </summary>
        /// <param name="operand1Formula">The new <see cref="Operand1Formula" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="CellValueConditionalFormattingRule" /> using the specified <paramref name="operand1Formula" /> for <see cref="Operand1Formula" /> and a deep clone of every other property.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002: DoNotExposeGenericLists")]
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords")]
        [SuppressMessage("Microsoft.Naming", "CA1719:ParameterNamesShouldNotMatchMemberNames")]
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames")]
        [SuppressMessage("Microsoft.Naming", "CA1722:IdentifiersShouldNotHaveIncorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration")]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly")]
        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public CellValueConditionalFormattingRule DeepCloneWithOperand1Formula(string operand1Formula)
        {
            var result = new CellValueConditionalFormattingRule
                             {
                                 Operator        = this.Operator,
                                 Operand1Formula = operand1Formula,
                                 Operand2Formula = this.Operand2Formula?.Clone().ToString(),
                                 StopIfTrue      = this.StopIfTrue,
                                 RangeStyle      = this.RangeStyle?.DeepClone(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="Operand2Formula" />.
        /// </summary>
        /// <param name="operand2Formula">The new <see cref="Operand2Formula" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="CellValueConditionalFormattingRule" /> using the specified <paramref name="operand2Formula" /> for <see cref="Operand2Formula" /> and a deep clone of every other property.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002: DoNotExposeGenericLists")]
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords")]
        [SuppressMessage("Microsoft.Naming", "CA1719:ParameterNamesShouldNotMatchMemberNames")]
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames")]
        [SuppressMessage("Microsoft.Naming", "CA1722:IdentifiersShouldNotHaveIncorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration")]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly")]
        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public CellValueConditionalFormattingRule DeepCloneWithOperand2Formula(string operand2Formula)
        {
            var result = new CellValueConditionalFormattingRule
                             {
                                 Operator        = this.Operator,
                                 Operand1Formula = this.Operand1Formula?.Clone().ToString(),
                                 Operand2Formula = operand2Formula,
                                 StopIfTrue      = this.StopIfTrue,
                                 RangeStyle      = this.RangeStyle?.DeepClone(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="StopIfTrue" />.
        /// </summary>
        /// <param name="stopIfTrue">The new <see cref="StopIfTrue" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="CellValueConditionalFormattingRule" /> using the specified <paramref name="stopIfTrue" /> for <see cref="StopIfTrue" /> and a deep clone of every other property.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002: DoNotExposeGenericLists")]
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords")]
        [SuppressMessage("Microsoft.Naming", "CA1719:ParameterNamesShouldNotMatchMemberNames")]
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames")]
        [SuppressMessage("Microsoft.Naming", "CA1722:IdentifiersShouldNotHaveIncorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration")]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly")]
        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public CellValueConditionalFormattingRule DeepCloneWithStopIfTrue(bool stopIfTrue)
        {
            var result = new CellValueConditionalFormattingRule
                             {
                                 Operator        = this.Operator,
                                 Operand1Formula = this.Operand1Formula?.Clone().ToString(),
                                 Operand2Formula = this.Operand2Formula?.Clone().ToString(),
                                 StopIfTrue      = stopIfTrue,
                                 RangeStyle      = this.RangeStyle?.DeepClone(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="RangeStyle" />.
        /// </summary>
        /// <param name="rangeStyle">The new <see cref="RangeStyle" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="CellValueConditionalFormattingRule" /> using the specified <paramref name="rangeStyle" /> for <see cref="RangeStyle" /> and a deep clone of every other property.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002: DoNotExposeGenericLists")]
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly")]
        [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix")]
        [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1716:IdentifiersShouldNotMatchKeywords")]
        [SuppressMessage("Microsoft.Naming", "CA1719:ParameterNamesShouldNotMatchMemberNames")]
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames")]
        [SuppressMessage("Microsoft.Naming", "CA1722:IdentifiersShouldNotHaveIncorrectPrefix")]
        [SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration")]
        [SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly")]
        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public CellValueConditionalFormattingRule DeepCloneWithRangeStyle(RangeStyle rangeStyle)
        {
            var result = new CellValueConditionalFormattingRule
                             {
                                 Operator        = this.Operator,
                                 Operand1Formula = this.Operand1Formula?.Clone().ToString(),
                                 Operand2Formula = this.Operand2Formula?.Clone().ToString(),
                                 StopIfTrue      = this.StopIfTrue,
                                 RangeStyle      = rangeStyle,
                             };

            return result;
        }

        /// <inheritdoc />
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        public override string ToString()
        {
            var result = Invariant($"OBeautifulCode.Excel.CellValueConditionalFormattingRule: Operator = {this.Operator.ToString() ?? "<null>"}, Operand1Formula = {this.Operand1Formula?.ToString(CultureInfo.InvariantCulture) ?? "<null>"}, Operand2Formula = {this.Operand2Formula?.ToString(CultureInfo.InvariantCulture) ?? "<null>"}, StopIfTrue = {this.StopIfTrue.ToString(CultureInfo.InvariantCulture) ?? "<null>"}, RangeStyle = {this.RangeStyle?.ToString() ?? "<null>"}.");

            return result;
        }
    }
}