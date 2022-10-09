using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.CommonUtils
{
    public static class ExcelUtils
    {
        #region DataTable List 相互转换

        /// <summary>
        /// Worksheet To DataTable 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowStart">行起点</param>
        /// <param name="ColEnd">列终点</param>
        /// <returns></returns>
        public static DataTable WorksheetToTable(ExcelWorksheet worksheet, int rowStart, int ColEnd)
        {
            int rows = worksheet.Dimension.End.Row;
            int cols = worksheet.Dimension.End.Column;
            DataTable dt = new DataTable(worksheet.Name);
            DataRow dr = null;
            bool isStop = false;
            ///行
            for (int i = 1; i < rows; i++)
            {
                if (i >= rowStart && !isStop)
                {
                    dr = dt.Rows.Add();
                }
                ///列
                for (int j = 1; j < cols; j++)
                {
                    if (i == rowStart - 1)
                    {
                        string value = GetString(worksheet.Cells[i, j].Value);
                        //string value = GetMegerValue(worksheet, i, j);
                        if (!dt.Columns.Contains(value))
                        {
                            dt.Columns.Add(value);
                        }
                    }
                    else
                    {
                        if (i < rowStart || j > ColEnd)
                        {
                            continue;
                        }
                        //string value = GetString(worksheet.Cells[i, j].Value);
                        string value = GetMegerValue(worksheet, i, j);
                        if (j == 1 && string.IsNullOrEmpty(value))
                        {
                            isStop = true;
                            dr.Delete();
                        }
                        dr[j - 1] = value;

                        //if (addColumns != null || addColumns.Count() != 0)
                        //{
                        //    foreach (var addItem in addColumns)
                        //    {
                        //        int columnIndex = addItem.Key;
                        //        string columnName = addItem.Value;
                        //        int currentColIndex = dt.Columns.IndexOf(columnName);
                        //        dr[currentColIndex] = GetString(worksheet.Cells[i, columnIndex].Value);
                        //    }
                        //}
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// List<T> To DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="array"></param>
        /// <returns></returns>
        public static DataTable IEnumerableToDataTable<T>(this IEnumerable<T> array)
        {
            var ret = new DataTable();

            foreach (PropertyDescriptor dp in TypeDescriptor.GetProperties(typeof(T)))
            {
                ///特性名称
                var attribute = dp.Attributes.OfType<DisplayNameAttribute>().FirstOrDefault();
                //属性名称
                //ret.Columns.Add(dp.Name);
                if (attribute != null)
                {
                    ret.Columns.Add(attribute.DisplayName);
                }
            }
            foreach (T item in array)
            {
                var Row = ret.NewRow();
                foreach (PropertyDescriptor dp in TypeDescriptor.GetProperties(typeof(T)))
                {
                    var attribute = dp.Attributes.OfType<DisplayNameAttribute>().FirstOrDefault();
                    if (attribute == null) { continue; }
                    Row[attribute.DisplayName] = dp.GetValue(item);
                }
                ret.Rows.Add(Row);
            }
            return ret;
        }

        /// <summary>
        /// Datable To List<T>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<T> ToModelList<T>(this DataTable dt) where T : new()
        {
            // 定义集合    
            List<T> ts = new List<T>();

            // 获得此模型的类型   
            Type type = typeof(T);
            string tempName = null, tempDescription = null;

            foreach (DataRow dr in dt.Rows)
            {
                bool isAdd = true;
                T t = new T();
                // 获得此模型的公共属性      
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    // 检查DataTable是否包含此列    
                    tempName = pi.Name;
                    tempDescription = pi == null ? null : ((DescriptionAttribute)Attribute.GetCustomAttribute(pi, typeof(DescriptionAttribute)))?.Description;
                    string column = tempDescription ?? tempName;
                    if (dt.Columns.Contains(column))
                    {
                        // 判断此属性是否有Setter      
                        if (!pi.CanWrite)
                            continue;
                        object value = dr[column];
                        try
                        {
                            if (value != DBNull.Value)
                            {
                                if (pi.PropertyType.ToString().Contains("System.Nullable"))
                                    value = Convert.ChangeType(value, Nullable.GetUnderlyingType(pi.PropertyType));
                                else
                                    value = Convert.ChangeType(value, pi.PropertyType);
                                pi.SetValue(t, value, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }
                if (isAdd)
                {
                    ts.Add(t);
                }
            }
            return ts;
        }

       
        /// <summary>
        /// 读取单元格数据
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string GetString(object obj)
        {
            try
            {
                if (obj == null)
                {
                    return "";
                }
                else
                {
                    return obj.ToString();
                }
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// 读取合并单元格内的数据
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static string GetMegerValue(ExcelWorksheet wSheet, int row, int column)
        {
            string range = wSheet.MergedCells[row, column];
            if (range == null)
                if (wSheet.Cells[row, column].Value != null)
                    return wSheet.Cells[row, column].Value.ToString();
                else
                    return "";
            object value =
                wSheet.Cells[(new ExcelAddress(range)).Start.Row, (new ExcelAddress(range)).Start.Column].Value;
            if (value != null)
                return value.ToString();
            else
                return "";
        }

        #endregion
    }
}
