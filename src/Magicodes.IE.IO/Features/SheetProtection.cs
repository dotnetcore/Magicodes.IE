namespace Magicodes.IE.IO
{
    /// <summary>
    /// Worksheet protection options. With the exception of <see cref="Sheet"/>, each property is <see langword="true"/> when the corresponding action is allowed.
    /// </summary>
    public sealed class SheetProtection
    {
        /// <summary>
        /// Gets or sets the password hash. This is a hash, not the plain-text password.
        /// </summary>
        public string? PasswordHash { get; init; }

        /// <summary>
        /// Gets or sets a value indicating whether the sheet is protected.
        /// </summary>
        public bool Sheet { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether formatting cells is allowed.
        /// </summary>
        public bool FormatCells { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether formatting columns is allowed.
        /// </summary>
        public bool FormatColumns { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether formatting rows is allowed.
        /// </summary>
        public bool FormatRows { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether inserting columns is allowed.
        /// </summary>
        public bool InsertColumns { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether inserting rows is allowed.
        /// </summary>
        public bool InsertRows { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether deleting columns is allowed.
        /// </summary>
        public bool DeleteColumns { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether deleting rows is allowed.
        /// </summary>
        public bool DeleteRows { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether sorting is allowed.
        /// </summary>
        public bool Sort { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether the auto-filter is allowed.
        /// </summary>
        public bool AutoFilter { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether selecting locked cells is allowed.
        /// </summary>
        public bool SelectLockedCells { get; init; } = false;

        /// <summary>
        /// Gets or sets a value indicating whether selecting unlocked cells is allowed.
        /// </summary>
        public bool SelectUnlockedCells { get; init; } = false;
    }
}
