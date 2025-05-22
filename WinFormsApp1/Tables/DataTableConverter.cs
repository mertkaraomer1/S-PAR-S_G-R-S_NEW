using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public static class DataTableConverter
    {
        public static List<T> ConvertTo<T>(DataTable datatable) where T : new()
        {
            List<T> Temp = new List<T>();
            try
            {
                List<string> columnsNames = new List<string>();
                foreach (DataColumn DataColumn in datatable.Columns)
                    columnsNames.Add(DataColumn.ColumnName);
                Temp = datatable.AsEnumerable().ToList().ConvertAll(row => getObject<T>(row, columnsNames));
                return Temp;
            }
            catch
            {
                return Temp;
            }

        }
        public static T getObject<T>(DataRow row, List<string> columnsName) where T : new()
        {
            T obj = new T();
            try
            {
                string columnname = "";
                string value = "";
                PropertyInfo[] Properties;
                Properties = typeof(T).GetProperties();
                foreach (PropertyInfo objProperty in Properties)
                {
                    columnname = columnsName.Find(name => name.ToLower() == GetDescriptionFromPropertyInfo(objProperty).ToLower());
                    if (!string.IsNullOrEmpty(columnname))
                    {
                        value = row[columnname].ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            objProperty.SetValue(obj, value, null);
                        }
                    }
                }
                return obj;
            }
            catch
            {
                return obj;
            }
        }
        public static string GetDescriptionFromPropertyInfo(PropertyInfo propertyInfo)
        {
            // PropertyInfo nesnesinden Description niteliği için kullanılan Attribute'ü al
            DescriptionAttribute descriptionAttribute = (DescriptionAttribute)Attribute.GetCustomAttribute(propertyInfo, typeof(DescriptionAttribute));

            // Description niteliği null değilse değerini döndür, aksi halde boş bir string döndür
            return descriptionAttribute != null ? descriptionAttribute.Description : string.Empty;
        }
    }
}
