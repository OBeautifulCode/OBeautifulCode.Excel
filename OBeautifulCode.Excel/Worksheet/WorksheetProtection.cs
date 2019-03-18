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
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The object to the right of the operator.</param>
        /// <returns>True if the two items are equal; false otherwise.</returns>
        public static bool operator ==(
            WorksheetProtection left,
            WorksheetProtection right)
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
                left.ClearTextPassword == right.ClearTextPassword;
            return result;
        }

        /// <summary>
        /// Determines whether two objects of type <see cref="WorksheetProtection"/> are not equal.
        /// </summary>
        /// <param name="left">The object to the left of the operator.</param>
        /// <param name="right">The item to compare.</param>
        /// <returns>True if the two items not equal; false otherwise.</returns>
        public static bool operator !=(
            WorksheetProtection left,
            WorksheetProtection right)
            => !(left == right);

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
        /// Creates a deep clone of this object.
        /// </summary>
        /// <returns>
        /// A deep clone of this object.
        /// </returns>
        public WorksheetProtection DeepClone()
        {
            var result = new WorksheetProtection
            {
                ClearTextPassword = this.ClearTextPassword,
            };

            return result;
        }
    }
}
