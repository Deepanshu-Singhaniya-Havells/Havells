using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Text;

namespace Havells.D365.Services.Utility
{
    public class Utilities
    {
        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {

                    if (pro.Name == column.ColumnName)
                    {
                        if (pro.PropertyType==typeof(string))
                        {
                            if (dr[column.ColumnName] != DBNull.Value)
                                pro.SetValue(obj, dr[column.ColumnName].ToString(), null);
                            else
                                pro.SetValue(obj, null, null);
                        }
                        else if (pro.PropertyType==typeof(double?))
                        {
                            if (dr[column.ColumnName] == DBNull.Value || dr[column.ColumnName].ToString() == "")
                                pro.SetValue(obj, null, null);
                            else
                                pro.SetValue(obj, Convert.ToDouble(dr[column.ColumnName]), null);
                        }
                        else if (pro.PropertyType==typeof(int?))
                        {

                            if (dr[column.ColumnName] == DBNull.Value || dr[column.ColumnName].ToString() == "")
                                pro.SetValue(obj, null,null);
                            else
                                pro.SetValue(obj, Convert.ToInt32(dr[column.ColumnName]), null);
                        }
                        else if (pro.PropertyType == typeof(long?))
                        {
                            if (dr[column.ColumnName] == DBNull.Value || dr[column.ColumnName].ToString() == "")
                                pro.SetValue(obj, null, null);
                            else
                                pro.SetValue(obj, Convert.ToInt64(dr[column.ColumnName]), null);
                        }
                        else if(pro.PropertyType==typeof(DateTime?))
                        {
                            if (dr[column.ColumnName] == DBNull.Value || dr[column.ColumnName].ToString() == "")
                                pro.SetValue(obj, null, null);
                            else
                                pro.SetValue(obj, Convert.ToDateTime(dr[column.ColumnName]), null);
                        }
                        else if (pro.PropertyType == typeof(bool?))
                        {
                            if (dr[column.ColumnName] == DBNull.Value || dr[column.ColumnName].ToString() == "")
                                pro.SetValue(obj, null, null);
                            else
                                pro.SetValue(obj, Convert.ToBoolean(dr[column.ColumnName]), null);
                        }
                        else
                            pro.SetValue(obj, dr[column.ColumnName], null);
                    }
                    else
                        continue;
                }
            }
            return obj;
        }

        public static string DataTableToJSON(DataTable table)
        {
            var JSONString = new StringBuilder();
            if (table.Rows.Count > 0)
            {
                JSONString.Append("[");
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    JSONString.Append("{");
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        if (j < table.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\",");
                        }
                        else if (j == table.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == table.Rows.Count - 1)
                    {
                        JSONString.Append("}");
                    }
                    else
                    {
                        JSONString.Append("},");
                    }
                }
                JSONString.Append("]");
            }
            return JSONString.ToString();
        }
    }
}
