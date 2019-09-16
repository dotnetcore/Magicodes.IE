using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models
{
    public enum ImporterProductType
    {
        [Display(Name = "第一")]
        One,

        [Display(Name = "第二")]
        Two
    }
}