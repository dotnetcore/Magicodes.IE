/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing;
using System;
using System.IO;
using System.Xml;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using Magicodes.IE.EPPlus;

namespace OfficeOpenXml
{
    /// <summary>
    /// An image that fills the background of the worksheet.
    /// </summary>
    public class ExcelBackgroundImage : XmlHelper
    {
        ExcelWorksheet _workSheet;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="nsm"></param>
        /// <param name="topNode">The topnode of the worksheet</param>
        /// <param name="workSheet">Worksheet reference</param>
        internal ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
            base(nsm, topNode)
        {
            _workSheet = workSheet;
        }

        const string BACKGROUNDPIC_PATH = "d:picture/@r:id";

        /// <summary>
        /// The background image of the worksheet. 
        /// The image will be saved internally as a jpg.
        /// </summary>
        public Image Image { get; private set; }

        internal IImageFormat ImageFormat { get; private set; }

        public void SetImage(byte[] imageBytes = null)
        {
            DeletePrevImage();
            if (imageBytes == null)
            {
                DeleteAllNode(BACKGROUNDPIC_PATH);
                return;
            }
            Image = Image.Load(imageBytes);
            //Image = Image.Load(imageBytes, out var imageFormat);
            using (MemoryStream imageStream = new MemoryStream(imageBytes))
            {
                var imageFormat = Image.GetImageFormat(imageStream);
                ImageFormat = imageFormat;
            }

            var imageInfo = _workSheet.Workbook._package.AddImage(imageBytes);
            var rel = _workSheet.Part.CreateRelationship(imageInfo.Uri, Packaging.TargetMode.Internal,
                ExcelPackage.schemaRelationships + "/image");
            SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
        }

        /// <summary>
        /// Set the picture from an image file. 
        /// The image file will be saved as a blob, so make sure Excel supports the image format.
        /// </summary>
        /// <param name="pictureFile">The image file.</param>
        public void SetFromFile(FileInfo pictureFile)
        {
            DeletePrevImage();

            try
            {
                var fileBytes = File.ReadAllBytes(pictureFile.FullName);
                //Image.Load(fileBytes, out var format);
                var format = Image.DetectFormat(fileBytes);

                string contentType = format.DefaultMimeType;
                var imageUri = XmlHelper.GetNewUri(_workSheet._package.Package,
                    "/xl/media/" +
                    pictureFile.Name.Substring(0, pictureFile.Name.Length - pictureFile.Extension.Length) + "{0}" +
                    pictureFile.Extension);

                var ii = _workSheet.Workbook._package.AddImage(fileBytes, imageUri, contentType);


                if (_workSheet.Part.Package.PartExists(imageUri) &&
                    ii.RefCount == 1) //The file exists with another content, overwrite it.
                {
                    //Remove the part if it exists
                    _workSheet.Part.Package.DeletePart(imageUri);
                }

                var imagePart = _workSheet.Part.Package.CreatePart(imageUri, contentType, CompressionLevel.None);
                //Save the picture to package.

                var stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
                stream.Write(fileBytes, 0, fileBytes.Length);

                var rel = _workSheet.Part.CreateRelationship(imageUri, Packaging.TargetMode.Internal,
                    ExcelPackage.schemaRelationships + "/image");
                SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("File is not a supported image-file or is corrupt", ex));
            }
        }

        private void DeletePrevImage()
        {
            var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
            if (relID != "")
            {
                var img = ImageCompat.GetImageAsByteArray(Image, ImageFormat);

                var ii = _workSheet.Workbook._package.GetImageInfo(img);

                //Delete the relation
                _workSheet.Part.DeleteRelationship(relID);

                //Delete the image if there are no other references.
                if (ii != null && ii.RefCount == 1)
                {
                    if (_workSheet.Part.Package.PartExists(ii.Uri))
                    {
                        _workSheet.Part.Package.DeletePart(ii.Uri);
                    }
                }
            }
        }
    }
}