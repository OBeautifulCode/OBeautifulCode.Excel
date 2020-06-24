// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellValueConditionalFormattingRule.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System.Diagnostics.CodeAnalysis;

    using OBeautifulCode.Type;

    /// <summary>
    /// Specifies a conditional formatting rule based on the value of a cell.
    /// </summary>
    public partial class CellValueConditionalFormattingRule : IModelViaCodeGen
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
    }
}
