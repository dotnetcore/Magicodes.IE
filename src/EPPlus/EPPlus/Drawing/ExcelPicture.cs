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
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;
using System.Xml;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Formats.Jpeg;
using Magicodes.IE.EPPlus;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing
    {
        #region "Constructors"

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlNode picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
            if (picNode != null)
            {
                RelPic = drawings.Part.GetRelationship(picNode.Attributes["r:embed"].Value);
                UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, RelPic.TargetUri);

                Part = drawings.Part.Package.GetPart(UriPic);
                // Image = Image.Load(Part.GetStream(), out var format);
                Image = Image.Load(Part.GetStream());
                var format = Image.GetImageFormat(Part.GetStream());
                ImageFormat = format;
                byte[] iby = ImageCompat.GetImageAsByteArray(Image, format);
                var ii = _drawings._package.LoadImage(iby, UriPic, Part);
                ImageHash = ii.Hash;

                //_height = _image.Height;
                //_width = _image.Width;
                string relID = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@r:id");
                if (!string.IsNullOrEmpty(relID))
                {
                    HypRel = drawings.Part.GetRelationship(relID);
                    _hyperlink = HypRel.TargetUri.IsAbsoluteUri ? new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri) : new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);

                    ((ExcelHyperLink)_hyperlink).ToolTip =
                        GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@tooltip");
                }
            }
        }

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image, IImageFormat format, Uri hyperlink) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode, node.SelectSingleNode("xdr:to", NameSpaceManager));
            _hyperlink = hyperlink;
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings),
                picNode);

            var package = drawings.Worksheet._package.Package;
            //Get the picture if it exists or save it if not.
            Image = image;
            ImageFormat = format;
            string relID = SavePicture(image, format);

            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
            _height = image.Height;
            _width = image.Width;
            SetPosDefaults(image);
            package.Flush();
        }

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, FileInfo imageFile, Uri hyperlink) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode, node.SelectSingleNode("xdr:to", NameSpaceManager));
            _hyperlink = hyperlink;
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings),
                picNode);

            //Changed to stream 2/4-13 (issue 14834). Thnx SClause
            var package = drawings.Worksheet._package.Package;
            using (var imageStream = new FileStream(imageFile.FullName, FileMode.Open, FileAccess.Read))
            {
                // Image = Image.Load(imageStream, out var format);
                Image = Image.Load(imageStream);
                var format = Image.GetImageFormat(imageStream);
                ImageFormat = format;

                var img = ImageCompat.GetImageAsByteArray(Image, format);
                UriPic = GetNewUri(package, "/xl/media/{0}" + imageFile.Name);
                var ii = _drawings._package.AddImage(img, UriPic, ContentType);
                string relID;
                if (!drawings._hashes.ContainsKey(ii.Hash))
                {
                    Part = ii.Part;
                    RelPic = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, ii.Uri),
                        Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                    relID = RelPic.Id;
                    _drawings._hashes.Add(ii.Hash, relID);
                    AddNewPicture(img, relID);
                }
                else
                {
                    relID = drawings._hashes[ii.Hash];
                    var rel = _drawings.Part.GetRelationship(relID);
                    UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                }

                ImageHash = ii.Hash;
                _height = Image.Height;
                _width = Image.Width;
                SetPosDefaults(Image);
                //Create relationship
                node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
                package.Flush();
            }
        }

        //Add a new image to the compare collection
        private void AddNewPicture(byte[] img, string relID)
        {
            var newPic = new ExcelDrawings.ImageCompare();
            newPic.image = img;
            newPic.relID = relID;
            //_drawings._pics.Add(newPic);
        }

        #endregion

        private string SavePicture(Image image, IImageFormat format)
        {
            byte[] img = ImageCompat.GetImageAsByteArray(image, format);

            var ii = _drawings._package.AddImage(img);


            ImageHash = ii.Hash;
            if (_drawings._hashes.ContainsKey(ii.Hash))
            {
                var relID = _drawings._hashes[ii.Hash];
                var rel = _drawings.Part.GetRelationship(relID);
                UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                return relID;
            }

            UriPic = ii.Uri;
            ImageHash = ii.Hash;

            //Set the Image and save it to the package.
            RelPic = _drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(_drawings.UriDrawing, UriPic),
                Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            //AddNewPicture(img, picRelation.Id);
            _drawings._hashes.Add(ii.Hash, RelPic.Id);
            return RelPic.Id;
        }

        private void SetPosDefaults(Image image)
        {
            EditAs = eEditAs.OneCell;
            SetPixelWidth(image.Width, (float)image.Metadata.HorizontalResolution);
            SetPixelHeight(image.Height, (float)image.Metadata.VerticalResolution);
        }

        private string PicStartXml()
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<xdr:nvPicPr>");

            if (_hyperlink == null)
            {
                xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
            }
            else
            {
                HypRel = _drawings.Part.CreateRelationship(_hyperlink, Packaging.TargetMode.External,
                    ExcelPackage.schemaHyperlink);
                xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\">", _id);
                if (HypRel != null)
                {
                    if (_hyperlink is ExcelHyperLink)
                    {
                        xml.AppendFormat(
                            "<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" tooltip=\"{1}\"/>",
                            HypRel.Id, ((ExcelHyperLink)_hyperlink).ToolTip);
                    }
                    else
                    {
                        xml.AppendFormat(
                            "<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" />",
                            HypRel.Id);
                    }
                }

                xml.Append("</xdr:cNvPr>");
            }

            xml.Append(
                "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

        internal string ImageHash { get; set; }

        /// <summary>
        /// The Image
        /// </summary>
        public Image Image { get; private set; }

        /// <summary>
        /// Image format
        /// If the picture is created from an Image this type is always Jpeg
        /// </summary>
        public IImageFormat ImageFormat { get; private set; }

        internal string ContentType => ImageFormat?.DefaultMimeType;

        /// <summary>
        /// Set the size of the image in percent from the orginal size
        /// Note that resizing columns / rows after using this function will effect the size of the picture
        /// </summary>
        /// <param name="Percent">Percent</param>
        public override void SetSize(int Percent)
        {
            if (Image == null)
            {
                base.SetSize(Percent);
            }
            else
            {
                _width = Image.Width;
                _height = Image.Height;

                _width = (int)(_width * ((decimal)Percent / 100));
                _height = (int)(_height * ((decimal)Percent / 100));

                SetPixelWidth(_width, (float) Image.Metadata.HorizontalResolution);
                SetPixelHeight(_height, (float) Image.Metadata.VerticalResolution);
            }
        }

        internal Uri UriPic { get; set; }
        internal Packaging.ZipPackageRelationship RelPic { get; set; }
        internal Packaging.ZipPackageRelationship HypRel { get; set; }
        internal Packaging.ZipPackagePart Part;

        internal new string Id
        {
            get { return Name; }
        }

        ExcelDrawingFill _fill = null;

        /// <summary>
        /// Fill
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr");
                }

                return _fill;
            }
        }

        ExcelDrawingBorder _border = null;

        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:ln");
                }

                return _border;
            }
        }

        private Uri _hyperlink = null;

        /// <summary>
        /// Hyperlink
        /// </summary>
        public Uri Hyperlink
        {
            get { return _hyperlink; }
        }

        internal override void DeleteMe()
        {
            _drawings._package.RemoveImage(ImageHash);
            base.DeleteMe();
        }

        public override void Dispose()
        {
            base.Dispose();
            _hyperlink = null;
            Image.Dispose();
            Image = null;
            ImageFormat = JpegFormat.Instance;
        }
    }
}