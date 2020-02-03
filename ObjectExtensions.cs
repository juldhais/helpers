using System;
using System.Linq;

namespace Helpers
{
    public static class ObjectExtensions
    {
        public static bool IsEmpty(this object value)
        {
            if (value == null) return true;
            if (string.IsNullOrWhiteSpace(value.ToString())) return true;

            return false;
        }

        public static bool IsNotEmpty(this object value)
        {
            return !IsEmpty(value);
        }

        public static bool IsGuidEmpty(this Guid value)
        {
            return value == Guid.Empty;
        }

        public static bool IsGuidEmpty(this Guid? value)
        {
            return (value == null || value == Guid.Empty);
        }

        public static bool IsGuidNotEmpty(this Guid value)
        {
            return !IsGuidEmpty(value);
        }

        public static bool IsGuidNotEmpty(this Guid? value)
        {
            return !IsGuidEmpty(value);
        }

        public static Guid? ToGuid(this object value)
        {
            try
            {
                return Guid.Parse(value.ToString());
            }
            catch
            {
                return null;
            }
        }

        public static int ToInteger(this object value)
        {
            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        public static decimal ToDecimal(this object value)
        {
            try
            {
                return Convert.ToDecimal(value);
            }
            catch
            {
                return 0;
            }
        }

        public static void MapFrom(this object model, object source)
        {
            var modelProperties = model.GetType().GetProperties();
            var sourceProperties = source.GetType().GetProperties();

            foreach (var sourceProperty in sourceProperties)
            {
                if (sourceProperty.GetCustomAttributes(true).Any(x => x.GetType().Name == "NotMappedToModelAttribute")) continue;

                if (modelProperties.Any(x => x.Name == sourceProperty.Name && x.PropertyType == sourceProperty.PropertyType))
                {
                    var value = sourceProperty.GetValue(source, null);

                    var modelProperty = modelProperties.First(x => x.Name == sourceProperty.Name && x.PropertyType == sourceProperty.PropertyType);

                    modelProperty.SetValue(model, value, null);
                }
            }
        }

        public static void CleanStringProperties(this object model)
        {
            var stringProperties = model.GetType().GetProperties().Where(x => x.PropertyType == typeof(string));

            foreach (var property in stringProperties)
            {
                var value = property.GetValue(model, null);
                if (value.IsEmpty()) property.SetValue(model, "");
            }
        }
    }

    public class NotMappedToModelAttribute : Attribute { }
}
