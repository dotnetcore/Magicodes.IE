using System;
using System.Globalization;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// The type of input a data-validation rule accepts.
    /// </summary>
    public enum DataValidationType
    {
        Any = 0,
        Integer = 1,
        Decimal = 2,
        List = 3,
        Date = 4,
        Time = 5,
        TextLength = 6,
        Custom = 7,
    }

    /// <summary>
    /// The comparison operator used by a data-validation rule.
    /// </summary>
    public enum DataValidationOperator
    {
        Between = 0,
        NotBetween = 1,
        Equal = 2,
        NotEqual = 3,
        LessThan = 4,
        LessThanOrEqual = 5,
        GreaterThan = 6,
        GreaterThanOrEqual = 7,
    }

    /// <summary>
    /// A data-validation rule applied to a range of cells on a worksheet.
    /// </summary>
    public sealed class DataValidation
    {
        /// <summary>
        /// Gets the cell range the rule applies to, for example <c>C2:C1000</c>.
        /// </summary>
        public string CellRange { get; }

        /// <summary>
        /// Gets the validation type.
        /// </summary>
        public DataValidationType Type { get; }

        /// <summary>
        /// Gets the comparison operator, or <see langword="null"/> for list validation.
        /// </summary>
        public DataValidationOperator? Operator { get; }

        /// <summary>
        /// Gets the first formula or value.
        /// </summary>
        public string? Formula1 { get; }

        /// <summary>
        /// Gets the second formula or value, used only for range comparisons.
        /// </summary>
        public string? Formula2 { get; }

        /// <summary>
        /// Gets a value indicating whether empty cells are allowed.
        /// </summary>
        public bool AllowBlank { get; }

        /// <summary>
        /// Gets a value indicating whether the input prompt is shown.
        /// </summary>
        public bool ShowInputMessage { get; }

        /// <summary>
        /// Gets a value indicating whether the error alert is shown.
        /// </summary>
        public bool ShowErrorMessage { get; }

        /// <summary>
        /// Gets the title of the input prompt.
        /// </summary>
        public string? PromptTitle { get; }

        /// <summary>
        /// Gets the body text of the input prompt.
        /// </summary>
        public string? PromptBody { get; }

        /// <summary>
        /// Creates a data-validation rule.
        /// </summary>
        public DataValidation(string cellRange, DataValidationType type, string? formula1, string? formula2 = null,
            DataValidationOperator? op = null, bool allowBlank = true, bool showInput = true, bool showError = true,
            string? promptTitle = null, string? promptBody = null)
        {
            CellRange = cellRange ?? throw new ArgumentNullException(nameof(cellRange));
            Type = type;
            Formula1 = formula1;
            Formula2 = formula2;
            Operator = op;
            AllowBlank = allowBlank;
            ShowInputMessage = showInput;
            ShowErrorMessage = showError;
            PromptTitle = promptTitle;
            PromptBody = promptBody;
        }

        /// <summary>
        /// Creates a list (drop-down) validation. <paramref name="listFormula"/> may be a quoted list or a cell range.
        /// </summary>
        public static DataValidation CreateList(string cellRange, string listFormula, bool allowBlank = true)
            => new(cellRange, DataValidationType.List, listFormula, null, op: null, allowBlank: allowBlank);

        /// <summary>
        /// Creates an integer range validation between <paramref name="min"/> and <paramref name="max"/>.
        /// </summary>
        public static DataValidation CreateInteger(string cellRange, int min, int max)
            => new(cellRange, DataValidationType.Integer, min.ToString(CultureInfo.InvariantCulture), max.ToString(CultureInfo.InvariantCulture), DataValidationOperator.Between);

        /// <summary>
        /// Creates a decimal range validation between <paramref name="min"/> and <paramref name="max"/>.
        /// </summary>
        public static DataValidation CreateDecimal(string cellRange, double min, double max)
            => new(cellRange, DataValidationType.Decimal, min.ToString(CultureInfo.InvariantCulture), max.ToString(CultureInfo.InvariantCulture), DataValidationOperator.Between);
    }
}
