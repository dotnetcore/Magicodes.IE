
using System;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// The anchor range for an image in a worksheet.
    /// </summary>
    public sealed class ImageAnchor
    {

        /// <summary>
        /// Gets the one-based picture number.
        /// </summary>
        public int ImageNo { get; }

        /// <summary>
        /// Gets the image file name.
        /// </summary>
        public string ImageName { get; }

        /// <summary>
        /// Gets the top-left cell of the anchor, for example <c>A1</c>.
        /// </summary>
        public string FromCell { get; }

        /// <summary>
        /// Gets the bottom-right cell of the anchor, for example <c>D8</c>.
        /// </summary>
        public string ToCell { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImageAnchor"/> class.
        /// </summary>
        public ImageAnchor(int imageNo, string imageName, string fromCell, string toCell)
        {
            if (imageName is null) throw new ArgumentNullException(nameof(imageName));
            if (fromCell is null) throw new ArgumentNullException(nameof(fromCell));
            if (toCell is null) throw new ArgumentNullException(nameof(toCell));
            ImageNo = imageNo;
            ImageName = imageName;
            FromCell = fromCell;
            ToCell = toCell;
        }
    }
}
