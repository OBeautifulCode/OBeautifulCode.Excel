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
    public partial class DocumentProperties : IModel<DocumentProperties>
    {
        /// <summary>
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are equal; otherwise false.</returns>
        public static bool operator ==(DocumentProperties left, DocumentProperties right)
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
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the equality operator.</param>
        /// <param name="right">The object to the right of the equality operator.</param>
        /// <returns>true if the two items are not equal; otherwise false.</returns>
        public static bool operator !=(DocumentProperties left, DocumentProperties right) => !(left == right);

        /// <inheritdoc />
        public bool Equals(DocumentProperties other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (ReferenceEquals(other, null))
            {
                return false;
            }

            var result = this.BuiltInDocumentPropertyKindToValueMap.IsEqualTo(other.BuiltInDocumentPropertyKindToValueMap);

            return result;
        }

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as DocumentProperties);

        /// <inheritdoc />
        public override int GetHashCode() => HashCodeHelper.Initialize()
            .Hash(this.BuiltInDocumentPropertyKindToValueMap)
            .Value;

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public DocumentProperties DeepClone()
        {
            var result = new DocumentProperties
                             {
                                 BuiltInDocumentPropertyKindToValueMap = this.BuiltInDocumentPropertyKindToValueMap?.ToDictionary(k => k.Key, v => v.Value?.Clone().ToString()),
                             };

            return result;
        }

        /// <summary>
        /// Deep clones this object with a new <see cref="BuiltInDocumentPropertyKindToValueMap" />.
        /// </summary>
        /// <param name="builtInDocumentPropertyKindToValueMap">The new <see cref="BuiltInDocumentPropertyKindToValueMap" />.  This object will NOT be deep cloned; it is used as-is.</param>
        /// <returns>New <see cref="DocumentProperties" /> using the specified <paramref name="builtInDocumentPropertyKindToValueMap" /> for <see cref="BuiltInDocumentPropertyKindToValueMap" /> and a deep clone of every other property.</returns>
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
        public DocumentProperties DeepCloneWithBuiltInDocumentPropertyKindToValueMap(IReadOnlyDictionary<BuiltInDocumentPropertyKind, string> builtInDocumentPropertyKindToValueMap)
        {
            var result = new DocumentProperties
                             {
                                 BuiltInDocumentPropertyKindToValueMap = builtInDocumentPropertyKindToValueMap,
                             };

            return result;
        }

        /// <inheritdoc />
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        public override string ToString()
        {
            var result = Invariant($"OBeautifulCode.Excel.DocumentProperties: BuiltInDocumentPropertyKindToValueMap = {this.BuiltInDocumentPropertyKindToValueMap?.ToString() ?? "<null>"}.");

            return result;
        }
    }
}