﻿// --------------------------------------------------------------------------------------------------------------------
// <auto-generated>
//   Generated using OBeautifulCode.CodeGen.ModelObject (1.0.86.0)
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
    public partial class WorksheetProtection : IModel<WorksheetProtection>
    {
        /// <summary>
        /// Determines whether two objects of type <see cref="WorksheetProtection"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are equal; otherwise false.</returns>
        public static bool operator ==(WorksheetProtection left, WorksheetProtection right)
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
        /// Determines whether two objects of type <see cref="WorksheetProtection"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are not equal; otherwise false.</returns>
        public static bool operator !=(WorksheetProtection left, WorksheetProtection right) => !(left == right);

        /// <inheritdoc />
        public bool Equals(WorksheetProtection other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (ReferenceEquals(other, null))
            {
                return false;
            }

            var result = this.ClearTextPassword.IsEqualTo(other.ClearTextPassword, StringComparer.Ordinal);

            return result;
        }

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as WorksheetProtection);

        /// <inheritdoc />
        public override int GetHashCode() => HashCodeHelper.Initialize()
            .Hash(this.ClearTextPassword)
            .Value;

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public WorksheetProtection DeepClone()
        {
            var result = new WorksheetProtection
                             {
                                 ClearTextPassword = this.ClearTextPassword?.Clone().ToString(),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="ClearTextPassword" />.
        /// </summary>
        /// <param name="clearTextPassword">The new <see cref="ClearTextPassword" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="WorksheetProtection" /> using the specified <paramref name="clearTextPassword" /> for <see cref="ClearTextPassword" /> and a deep clone of every other property.</returns>
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
        public WorksheetProtection DeepCloneWithClearTextPassword(string clearTextPassword)
        {
            var result = new WorksheetProtection
                             {
                                 ClearTextPassword = clearTextPassword,
                             };

            return result;
        }

        /// <inheritdoc />
        public override string ToString()
        {
            var result = Invariant($"OBeautifulCode.Excel.WorksheetProtection: ClearTextPassword = {this.ClearTextPassword?.ToString(CultureInfo.InvariantCulture) ?? "<null>"}.");

            return result;
        }
    }
}