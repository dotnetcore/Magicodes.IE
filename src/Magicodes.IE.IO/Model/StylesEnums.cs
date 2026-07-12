using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// The border style applied to a cell or range.
    /// </summary>
    public enum BorderStyle
    {
        /// <summary>
        /// No border.
        /// </summary>
        None = 0,

        /// <summary>
        /// A thin border.
        /// </summary>
        Thin = 1,

        /// <summary>
        /// A medium border.
        /// </summary>
        Medium = 2,

        /// <summary>
        /// A thick border.
        /// </summary>
        Thick = 3,

        /// <summary>
        /// A dashed border.
        /// </summary>
        Dashed = 4,

        /// <summary>
        /// A dotted border.
        /// </summary>
        Dotted = 5,

        /// <summary>
        /// A double border.
        /// </summary>
        Double = 6,
    }

    /// <summary>
    /// The vertical alignment of content within a cell.
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>
        /// No explicit alignment; Excel's default is used.
        /// </summary>
        None = 0,

        /// <summary>
        /// Align content to the top of the cell.
        /// </summary>
        Top = 1,

        /// <summary>
        /// Center content vertically.
        /// </summary>
        Center = 2,

        /// <summary>
        /// Align content to the bottom of the cell.
        /// </summary>
        Bottom = 3,
    }
}
