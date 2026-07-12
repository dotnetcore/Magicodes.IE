using System;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// The comparison operator for a conditional-formatting rule.
    /// </summary>
    public enum CfOperator
    {
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        Between,
        NotBetween,
    }

    /// <summary>
    /// A single conditional-formatting rule.
    /// </summary>
    public sealed class CfRule
    {
        /// <summary>
        /// Gets the comparison operator.
        /// </summary>
        public CfOperator Operator { get; }

        /// <summary>
        /// Gets the first formula or value.
        /// </summary>
        public string Formula1 { get; }

        /// <summary>
        /// Gets the second formula or value, used only for range comparisons.
        /// </summary>
        public string? Formula2 { get; }

        /// <summary>
        /// Creates a conditional-formatting rule.
        /// </summary>
        public CfRule(CfOperator op, string formula1, string? formula2 = null)
        {
            Operator = op;
            Formula1 = formula1 ?? throw new ArgumentNullException(nameof(formula1));
            Formula2 = formula2;
        }
    }

    /// <summary>
    /// Conditional formatting applied to a cell range.
    /// </summary>
    public sealed class ConditionalFormatting
    {
        /// <summary>
        /// Gets the cell range the rules apply to.
        /// </summary>
        public string CellRange { get; }

        /// <summary>
        /// Gets the rules, evaluated in order.
        /// </summary>
        public CfRule[] Rules { get; }

        /// <summary>
        /// Creates a set of conditional-formatting rules for the specified range.
        /// </summary>
        public ConditionalFormatting(string cellRange, params CfRule[] rules)
        {
            CellRange = cellRange ?? throw new ArgumentNullException(nameof(cellRange));
            Rules = rules ?? throw new ArgumentNullException(nameof(rules));
        }
    }
}
