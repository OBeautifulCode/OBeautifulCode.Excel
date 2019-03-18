// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocumentProperties.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using OBeautifulCode.Math.Recipes;
    using OBeautifulCode.Type;

    /// <summary>
    /// Specifies the properties of a workbook document.
    /// </summary>
    public class DocumentProperties : IEquatable<DocumentProperties>, IDeepCloneable<DocumentProperties>
    {
        /// <summary>
        /// Gets or sets a map of property kind to it's value.
        /// </summary>
        public IReadOnlyDictionary<BuiltInDocumentPropertyKind, string> BuiltInDocumentPropertyKindToValueMap { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            DocumentProperties left,
            DocumentProperties right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
            {
                return false;
            }

            if ((left.BuiltInDocumentPropertyKindToValueMap == null) ||
                (right.BuiltInDocumentPropertyKindToValueMap == null))
            {
                if ((left.BuiltInDocumentPropertyKindToValueMap == null) &&
                    (right.BuiltInDocumentPropertyKindToValueMap == null))
                {
                    return true;
                }

                return false;
            }

            var leftProperties = left.BuiltInDocumentPropertyKindToValueMap.OrderBy(_ => _.Key).ToList();
            var rightProperties = right.BuiltInDocumentPropertyKindToValueMap.OrderBy(_ => _.Key).ToList();

            var result = leftProperties.Select(_ => _.Key).SequenceEqual(rightProperties.Select(_ => _.Key)) &&
                         leftProperties.Select(_ => _.Value).SequenceEqual(rightProperties.Select(_ => _.Value));

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            DocumentProperties left,
            DocumentProperties right)
            => !(left == right);

        /// <inheritdoc />
        public bool Equals(DocumentProperties other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as DocumentProperties);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .HashElements(this.BuiltInDocumentPropertyKindToValueMap?.OrderBy(_ => _.Key).Select(_ => _.Key))
                .HashElements(this.BuiltInDocumentPropertyKindToValueMap?.OrderBy(_ => _.Value).Select(_ => _.Value))
                .Value;

        /// <inheritdoc />
        public object Clone() => this.DeepClone();

        /// <inheritdoc />
        public DocumentProperties DeepClone()
        {
            var result = new DocumentProperties
            {
                BuiltInDocumentPropertyKindToValueMap = this.BuiltInDocumentPropertyKindToValueMap?.ToDictionary(_ => _.Key, _ => _.Value),
            };

            return result;
        }
    }
}
