// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetProtection.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    using OBeautifulCode.Math.Recipes;

    /// <summary>
    /// The worksheet protection configuration.
    /// </summary>
    public class WorksheetProtection : IEquatable<WorksheetProtection>
    {
        /// <summary>
        /// Gets or sets the clear text password.
        /// </summary>
        public string ClearTextPassword { get; set; }

        /// <summary>
        /// Determines whether two objects of type <see cref="WorksheetProtection"/> are equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The second item to compare.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            WorksheetProtection item1,
            WorksheetProtection item2)
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
                item1.ClearTextPassword == item2.ClearTextPassword;
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="WorksheetProtection"/> are not equal.
        /// </summary>
        /// <param name="item1">The first item to compare.</param>
        /// <param name="item2">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            WorksheetProtection item1,
            WorksheetProtection item2)
            => !(item1 == item2);

        /// <inheritdoc />
        public bool Equals(WorksheetProtection other) => this == other;

        /// <inheritdoc />
        public override bool Equals(object obj) => this == (obj as WorksheetProtection);

        /// <inheritdoc />
        public override int GetHashCode() =>
            HashCodeHelper.Initialize()
                .Hash(this.ClearTextPassword)
                .Value;

        /// <summary>
        /// Creates a clone of this object.
        /// </summary>
        /// <returns>
        /// A clone of this object.
        /// </returns>
        public WorksheetProtection Clone()
        {
            var result = new WorksheetProtection
            {
                ClearTextPassword = this.ClearTextPassword,
            };

            return result;
        }
    }
}
