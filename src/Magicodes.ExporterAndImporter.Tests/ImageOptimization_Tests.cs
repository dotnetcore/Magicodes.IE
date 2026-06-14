using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Shouldly;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ImageOptimization_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public ImageOptimization_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        #region IdentifyImage

        [Fact(DisplayName = "IdentifyImage_PNG图片_返回正确元数据")]
        public void IdentifyImage_Png_ReturnsCorrectMetadata()
        {
            var pngPath = Path.Combine("TestFiles", "ExporterTest.png");
            var bytes = File.ReadAllBytes(pngPath);

            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            info.Width.ShouldBeGreaterThan(0);
            info.Height.ShouldBeGreaterThan(0);
            info.ContentType.ShouldBe("image/png");
            _output.WriteLine($"PNG: {info.Width}x{info.Height}, ContentType={info.ContentType}");
        }

        [Fact(DisplayName = "IdentifyImage_JPEG图片_返回正确元数据")]
        public void IdentifyImage_Jpeg_ReturnsCorrectMetadata()
        {
            var jpegPath = Path.Combine("TestFiles", "Images", "4.Jpeg");
            var bytes = File.ReadAllBytes(jpegPath);

            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            info.Width.ShouldBe(321);
            info.Height.ShouldBe(201);
            info.ContentType.ShouldBe("image/jpeg");
            _output.WriteLine($"JPEG: {info.Width}x{info.Height}, ContentType={info.ContentType}");
        }

        [Fact(DisplayName = "IdentifyImage_ZeroDPI图片_正确读取尺寸")]
        public void IdentifyImage_ZeroDpi_CorrectSize()
        {
            var zeroDpiPath = Path.Combine("TestFiles", "Images", "zero-DPI.Jpeg");
            var bytes = File.ReadAllBytes(zeroDpiPath);

            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            info.Width.ShouldBe(155);
            info.Height.ShouldBe(155);
            info.ContentType.ShouldBe("image/jpeg");
            _output.WriteLine($"Zero-DPI: {info.Width}x{info.Height}, ContentType={info.ContentType}");
        }

        [Fact(DisplayName = "IdentifyImage_无效字节_抛出异常")]
        public void IdentifyImage_InvalidBytes_ThrowsException()
        {
            var invalidBytes = new byte[] { 0x00, 0x01, 0x02, 0x03 };

            Should.Throw<InvalidOperationException>(() =>
                Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(invalidBytes));
        }

        [Fact(DisplayName = "IdentifyImage_空数组_抛出异常")]
        public void IdentifyImage_EmptyArray_ThrowsException()
        {
            Should.Throw<Exception>(() =>
                Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(Array.Empty<byte>()));
        }

        #endregion

        #region DecodeBase64ToBytes

        [Fact(DisplayName = "DecodeBase64ToBytes_有效Base64_正确解码")]
        public void DecodeBase64ToBytes_ValidBase64_ReturnsCorrectBytes()
        {
            var original = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            var base64 = Convert.ToBase64String(original);

            var result = Magicodes.IE.Excel.Images.ImageExtensions.DecodeBase64ToBytes(base64);

            result.ShouldBe(original);
        }

        [Fact(DisplayName = "DecodeBase64ToBytes_带空白字符_正确去除并解码")]
        public void DecodeBase64ToBytes_WithWhitespace_StripsAndDecodes()
        {
            var original = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
            var base64 = Convert.ToBase64String(original);
            var base64WithWhitespace = base64.Insert(2, "\r\n").Insert(5, " ");

            var result = Magicodes.IE.Excel.Images.ImageExtensions.DecodeBase64ToBytes(base64WithWhitespace);

            result.ShouldBe(original);
        }

        #endregion

        #region ReadImageBytes

        [Fact(DisplayName = "ReadImageBytes_有效文件_返回正确字节")]
        public void ReadImageBytes_ValidFile_ReturnsCorrectBytes()
        {
            var path = Path.Combine("TestFiles", "Images", "4.Jpeg");
            var expected = File.ReadAllBytes(path);

            var result = Magicodes.IE.Excel.Images.ImageExtensions.ReadImageBytes(path);

            result.ShouldBe(expected);
        }

        [Fact(DisplayName = "ReadImageBytes_不存在文件_抛出异常")]
        public void ReadImageBytes_NonExistentPath_ThrowsException()
        {
            Should.Throw<FileNotFoundException>(() =>
                Magicodes.IE.Excel.Images.ImageExtensions.ReadImageBytes("nonexistent.jpg"));
        }

        #endregion

        #region AddPictureFromBytes

        [Fact(DisplayName = "AddPictureFromBytes_PNG图片_正确添加Drawing")]
        public void AddPictureFromBytes_Png_AddsDrawingCorrectly()
        {
            var pngPath = Path.Combine("TestFiles", "ExporterTest.png");
            var bytes = File.ReadAllBytes(pngPath);
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            var filePath = GetTestFilePath($"{nameof(AddPictureFromBytes_Png_AddsDrawingCorrectly)}.xlsx");
            DeleteFile(filePath);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                var pic = sheet.Drawings.AddPictureFromBytes("test.png", bytes, info.ContentType, info.Width, info.Height);
                pic.From.Column = 0;
                pic.From.Row = 0;
                pic.SetSize(info.Width * 7, info.Height);

                package.SaveAs(new FileInfo(filePath));
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(1);
                var pic = (ExcelPicture)sheet.Drawings[0];
                pic.From.Column.ShouldBe(0);
                pic.From.Row.ShouldBe(0);
            }
        }

        [Fact(DisplayName = "AddPictureFromBytes_JPEG图片_正确添加Drawing")]
        public void AddPictureFromBytes_Jpeg_AddsDrawingCorrectly()
        {
            var jpegPath = Path.Combine("TestFiles", "Images", "4.Jpeg");
            var bytes = File.ReadAllBytes(jpegPath);
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            var filePath = GetTestFilePath($"{nameof(AddPictureFromBytes_Jpeg_AddsDrawingCorrectly)}.xlsx");
            DeleteFile(filePath);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                var pic = sheet.Drawings.AddPictureFromBytes("test.jpg", bytes, info.ContentType, info.Width, info.Height);
                pic.From.Column = 0;
                pic.From.Row = 0;

                package.SaveAs(new FileInfo(filePath));
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(1);
            }
        }

        [Fact(DisplayName = "AddPictureFromBytes_重复名称_抛出异常")]
        public void AddPictureFromBytes_DuplicateName_ThrowsException()
        {
            var bytes = File.ReadAllBytes(Path.Combine("TestFiles", "Images", "4.Jpeg"));
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Drawings.AddPictureFromBytes("same.jpg", bytes, info.ContentType, info.Width, info.Height);

                Should.Throw<Exception>(() =>
                    sheet.Drawings.AddPictureFromBytes("same.jpg", bytes, info.ContentType, info.Width, info.Height));
            }
        }

        [Fact(DisplayName = "AddPictureFromBytes_空字节_抛出异常")]
        public void AddPictureFromBytes_EmptyBytes_ThrowsException()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");

                Should.Throw<ArgumentException>(() =>
                    sheet.Drawings.AddPictureFromBytes("test.jpg", Array.Empty<byte>(), "image/jpeg", 100, 100));
            }
        }

        [Fact(DisplayName = "AddPictureFromBytes_Null字节_抛出异常")]
        public void AddPictureFromBytes_NullBytes_ThrowsException()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");

                Should.Throw<ArgumentException>(() =>
                    sheet.Drawings.AddPictureFromBytes("test.jpg", null, "image/jpeg", 100, 100));
            }
        }

        [Fact(DisplayName = "AddPictureFromBytes_相同图片_去重存储")]
        public void AddPictureFromBytes_SameImageTwice_Deduplicates()
        {
            var bytes = File.ReadAllBytes(Path.Combine("TestFiles", "Images", "4.Jpeg"));
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            var filePath = GetTestFilePath($"{nameof(AddPictureFromBytes_SameImageTwice_Deduplicates)}.xlsx");
            DeleteFile(filePath);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Drawings.AddPictureFromBytes("img1.jpg", bytes, info.ContentType, info.Width, info.Height);
                sheet.Drawings.AddPictureFromBytes("img2.jpg", bytes, info.ContentType, info.Width, info.Height);

                package.SaveAs(new FileInfo(filePath));
            }

            // 验证去重：两个 drawing 都存在
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(2);
            }

            // 验证去重：对比单图片文件大小，双图片文件不应是单图片的 2 倍
            var singleFilePath = GetTestFilePath($"{nameof(AddPictureFromBytes_SameImageTwice_Deduplicates)}_single.xlsx");
            DeleteFile(singleFilePath);
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Drawings.AddPictureFromBytes("img1.jpg", bytes, info.ContentType, info.Width, info.Height);
                package.SaveAs(new FileInfo(singleFilePath));
            }

            var singleSize = new FileInfo(singleFilePath).Length;
            var doubleSize = new FileInfo(filePath).Length;
            // 去重后，双图片文件大小应小于单图片的 1.8 倍（非 2 倍）
            doubleSize.ShouldBeLessThan((long)(singleSize * 1.8));
            _output.WriteLine($"单图片: {singleSize} bytes, 双图片(去重): {doubleSize} bytes, 比值: {(double)doubleSize/singleSize:F2}");
        }

        #endregion

        #region ExcelPicture 原始字节

        [Fact(DisplayName = "ExcelPicture_原始字节构造_Image属性为Null")]
        public void ExcelPicture_FromRawBytes_ImagePropertyIsNull()
        {
            var bytes = File.ReadAllBytes(Path.Combine("TestFiles", "Images", "4.Jpeg"));
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                using (var pic = sheet.Drawings.AddPictureFromBytes("test.jpg", bytes, info.ContentType, info.Width, info.Height))
                {
                    pic.Image.ShouldBeNull();
                    pic.ShouldNotBeNull();
                }
            }
        }

        [Fact(DisplayName = "ExcelPicture_原始字节构造_Dispose不抛异常")]
        public void ExcelPicture_FromRawBytes_DisposeDoesNotThrow()
        {
            var bytes = File.ReadAllBytes(Path.Combine("TestFiles", "Images", "4.Jpeg"));
            var info = Magicodes.IE.Excel.Images.ImageExtensions.IdentifyImage(bytes);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                var pic = sheet.Drawings.AddPictureFromBytes("test.jpg", bytes, info.ContentType, info.Width, info.Height);
                Should.NotThrow(() => pic.Dispose());
            }
        }

        #endregion

        #region 端到端导出

        [Fact(DisplayName = "ExportPicture_多行本地图片_并行加载全部正确导出")]
        public async Task ExportPicture_MultipleLocalImages_AllExported()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPicture_MultipleLocalImages_AllExported)}.xlsx");
            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithPicture>(20);
            var url = Path.Combine("TestFiles", "ExporterTest.png");
            foreach (var item in data)
            {
                item.Img1 = url;
                item.Img = url;
            }

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                // 20行 × 2列图片 = 40 个 drawing
                sheet.Drawings.Count.ShouldBe(40);
                _output.WriteLine($"成功导出 {sheet.Drawings.Count} 个图片");
            }
        }

        [Fact(DisplayName = "ExportPicture_混合有效无效URL_有效导出无效显示Alt")]
        public async Task ExportPicture_MixValidAndInvalid_ExportedCorrectly()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPicture_MixValidAndInvalid_ExportedCorrectly)}.xlsx");
            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithPicture>(5);
            var validUrl = Path.Combine("TestFiles", "ExporterTest.png");
            data[0].Img1 = validUrl;  // 有效
            data[1].Img1 = "nonexistent.jpg";  // 无效
            data[2].Img1 = validUrl;  // 有效
            data[3].Img1 = null;  // null
            data[4].Img1 = validUrl;  // 有效
            foreach (var item in data) item.Img = null;

            var result = await exporter.Export(filePath, data);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                // 3 个有效图片 (data[0], data[2], data[4])
                var validDrawings = sheet.Drawings.Count;
                validDrawings.ShouldBe(3);
                // 无效的应显示 Alt 文本 "404"
                sheet.Cells["G2"].Value.ShouldBe("404"); // data[1].Img1 无效
                _output.WriteLine($"有效图片: {validDrawings}, 无效行显示 Alt 文本");
            }
        }

        [Fact(DisplayName = "ExportPicture_Base64图片_正确导出")]
        public async Task ExportPicture_Base64Images_ExportedCorrectly()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPicture_Base64Images_ExportedCorrectly)}.xlsx");
            DeleteFile(filePath);

            var pngBytes = File.ReadAllBytes(Path.Combine("TestFiles", "ExporterTest.png"));
            var base64 = Convert.ToBase64String(pngBytes);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithPicture>(3);
            foreach (var item in data)
            {
                item.Img1 = base64;
                item.Img = null;
            }

            var result = await exporter.Export(filePath, data);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(3);
                _output.WriteLine($"Base64 图片导出: {sheet.Drawings.Count} 个");
            }
        }

        [Fact(DisplayName = "ExportPicture_相同URL_去重后全部正确导出")]
        public async Task ExportPicture_DuplicateUrls_AllExported()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPicture_DuplicateUrls_AllExported)}.xlsx");
            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithPicture>(10);
            var sameUrl = Path.Combine("TestFiles", "ExporterTest.png");
            foreach (var item in data)
            {
                item.Img1 = sameUrl; // 全部使用同一 URL
                item.Img = null;
            }

            var result = await exporter.Export(filePath, data);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                // 10 个 drawing（每行一个）
                sheet.Drawings.Count.ShouldBe(10);
                _output.WriteLine($"10 行相同 URL → {sheet.Drawings.Count} 个 drawing");
            }
        }

        #endregion
    }
}
