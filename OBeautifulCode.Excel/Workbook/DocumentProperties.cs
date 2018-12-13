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

    /// <summary>
    /// Specifies the properties of a workbook document.
    /// </summary>
    public class DocumentProperties : IEquatable<DocumentProperties>
    {
        /// <summary>
        /// Gets or sets a map of property kind to it's value.
        /// </summary>
        public IReadOnlyDictionary<BuiltInDocumentPropertyKind, string> BuiltInDocumentPropertyKindToValueMap { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            DocumentProperties item1,
            DocumentProperties item2)
        {
            if (ReferenceEquals(item1, item2))
            {
                return true;
            }

            if (ReferenceEquals(item1, null) || ReferenceEquals(item2, null))
            {
                return false;
            }

            if ((item1.BuiltInDocumentPropertyKindToValueMap == null) ||
                (item2.BuiltInDocumentPropertyKindToValueMap == null))
            {
                if ((item1.BuiltInDocumentPropertyKindToValueMap == null) &&
                    (item2.BuiltInDocumentPropertyKindToValueMap == null))
                {
                    return true;
                }

                return false;
            }

            var item1Properties = item1.BuiltInDocumentPropertyKindToValueMap.OrderBy(_ => _.Key).ToList();
            var item2Properties = item2.BuiltInDocumentPropertyKindToValueMap.OrderBy(_ => _.Key).ToList();

            var result = item1Properties.Select(_ => _.Key).SequenceEqual(item2Properties.Select(_ => _.Key)) &&
                         item1Properties.Select(_ => _.Value).SequenceEqual(item2Properties.Select(_ => _.Value));

            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="DocumentProperties"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            DocumentProperties item1,
            DocumentProperties item2)
            => !(item1 == item2);

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

        /// <summary>
        /// Creates a clone of this object.
        /// </summary>
        /// <returns>
        /// A clone of this object.
        /// </returns>
        public DocumentProperties Clone()
        {
            var result = new DocumentProperties
            {
                BuiltInDocumentPropertyKindToValueMap = this.BuiltInDocumentPropertyKindToValueMap?.ToDictionary(_ => _.Key, _ => _.Value),
            };

            return result;
        }
    }
}
