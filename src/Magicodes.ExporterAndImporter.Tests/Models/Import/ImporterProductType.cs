// ======================================================================
// 
//           filename : ImporterProductType.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public enum ImporterProductType
    {
        [Display(Name = "第一")] One,
        [Display(Name = "第二")] Two
    }
}