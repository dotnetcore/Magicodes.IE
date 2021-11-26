using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace Magicodes.IE.Core
{
    public class DynamicStringLengthAttribute : StringLengthAttribute
    {
        private static readonly FieldInfo MaximumLengthField;

        static DynamicStringLengthAttribute()
        {
            MaximumLengthField = typeof(StringLengthAttribute).GetField(
                "<MaximumLength>k__BackingField",
                BindingFlags.Instance | BindingFlags.NonPublic
            );
            Debug.Assert(MaximumLengthField != null, nameof(MaximumLengthField) + " != null");
        }

        /// <param name="sourceType">A type to get the values of the properties</param>
        /// <param name="maximumLengthPropertyName">The name of the public static property for the <see cref="StringLengthAttribute.MaximumLength"/></param>
        /// <param name="minimumLengthPropertyName">The name of the public static property for the <see cref="StringLengthAttribute.MinimumLength"/></param>
        public DynamicStringLengthAttribute(
            Type sourceType,
            string maximumLengthPropertyName,
            string minimumLengthPropertyName = null)
            : base(0)
        {
            if (sourceType is null)
            {
                throw new ArgumentNullException(nameof(sourceType));
            }

            if (maximumLengthPropertyName != null)
            {
                var maximumLengthProperty = sourceType.GetProperty(
                    maximumLengthPropertyName,
                    BindingFlags.Static | BindingFlags.Public
                );
                Debug.Assert(maximumLengthProperty != null, nameof(maximumLengthProperty) + " != null");
                MaximumLengthField.SetValue(this, (int)maximumLengthProperty.GetValue(null));
            }

            if (minimumLengthPropertyName != null)
            {
                var minimumLengthProperty = sourceType.GetProperty(
                    minimumLengthPropertyName,
                    BindingFlags.Static | BindingFlags.Public
                );
                Debug.Assert(minimumLengthProperty != null, nameof(minimumLengthProperty) + " != null");
                MinimumLength = (int)minimumLengthProperty.GetValue(null);
            }
        }
    }
}
