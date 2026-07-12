
using System;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// A cell comment in a worksheet.
    /// </summary>
    public sealed class Comment
    {

        /// <summary>
        /// Gets the zero-based row of the cell.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the zero-based column of the cell.
        /// </summary>
        public int Col { get; }

        /// <summary>
        /// Gets the author of the comment.
        /// </summary>
        public string Author { get; }

        /// <summary>
        /// Gets the comment text.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Comment"/> class.
        /// </summary>
        public Comment(int row, int col, string author, string text)
        {
            if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (col < 0) throw new ArgumentOutOfRangeException(nameof(col));
            Row = row;
            Col = col;
            Author = author ?? throw new ArgumentNullException(nameof(author));
            Text = text ?? throw new ArgumentNullException(nameof(text));
        }
    }
}
