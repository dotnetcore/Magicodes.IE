using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// 数据验证帮助类
    /// </summary>
    public static class ValidatorHelper
    {
        /// <summary>
        /// 校验数据合法性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="validationResults"></param>
        /// <returns></returns>
        public static bool TryValidate(object obj, out ICollection<ValidationResult> validationResults)
        {
            var context = new ValidationContext(obj, null, null);
            validationResults = new List<ValidationResult>();

            return Validator.TryValidateObject(obj, context, validationResults, true);
        }
    }
}