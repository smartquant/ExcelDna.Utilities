/*
The MIT License (MIT)

Copyright (c) 2014 Joachim Loebb

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExcelDna.Utilities
{
    public static class DataTableEx
    {

        public static object[,] ToVariant(this DataTable dt, bool header = false)
        {
            if (dt.Columns.Count == 0) return new object[1, 1] { { 0 } };

            int n = dt.Rows.Count, cols = dt.Columns.Count;
            int rows = header ? n + 1 : n;
            int start = header ? 1 : 0;


            object[,] retval = new object[rows, cols];

            if (header)
            {
                for (int i = 0; i < cols; i++)
                    retval[0, i] = dt.Columns[i].Caption;
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < cols; j++)
                    retval[start, j] = dt.Rows[i][dt.Columns[j].ColumnName];
                start++;
            }

            return retval;

        }

        public static String ToString(this DataTable dt, string sep = ",", bool header = false) 
        {
            var sb = new StringBuilder();
            if (dt.Rows.Count == 0) return string.Empty;

            int n = dt.Rows.Count, cols = dt.Columns.Count;
            int rows = header ? n + 1 : n;
            int start = header ? 1 : 0;


            if (header)
            {
                for (int i = 0; i < cols; i++)
                    sb.Append(dt.Columns[i].Caption + sep);
                sb.Length -= sep.Length;
                sb.AppendLine();
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < cols; j++)
                    sb.Append(dt.Rows[i][dt.Columns[j].ColumnName].ToString() + sep);

                sb.Length -= sep.Length;
                sb.AppendLine();
                start++;
            }
            return sb.ToString();
        }

        public static void AddRange(this DataTable dt, object vt)
        {
            object[,] vtdata = vt as object[,];
            if (vtdata == null) throw new ArgumentException("vt must be an variant array!");

            int n = vtdata.GetLength(0);
            int k = vtdata.GetLength(1);

            var cols = dt.Columns;
            if (cols.Count != k)
                throw new ArgumentException("Number of columns does not match the columns in vt!");

            for (int i = 0; i < n; i++)
            {
                var row = dt.NewRow();
                for (int j = 0; j < k; j++)
                {
                    row[cols[j].ColumnName] = vtdata[i, j].ConvertTo(cols[j].DataType);
                }
                dt.Rows.Add(row);
            }
        }

        public static string GetHeader(this DataTable dt, string sep = ",")
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                 sb.Append(dt.Columns[i].Caption + sep);
            }
            return sb.ToString(0, sb.Length - sep.Length);
        }

        public static DataTable CreateDataTable(this object vt, bool header = true)
        {
            
            object[,] vtdata = vt as object[,];
            if (vtdata == null) throw new ArgumentException("vt must be a 2-dimensional variant array!");

            var dt = new DataTable();

            int n = vtdata.GetLength(0);
            int k = vtdata.GetLength(1);
            int start = (header) ? 1 : 0;

            if (header)
            {
                for (int j = 0; j < k; j++)
                {
                    var col = vtdata[0, j].ToString();
                    dt.Columns.Add(new DataColumn(col, vtdata[1,j].GetType()));
                }
            }
            else
            {
                for (int j = 0; j < k; j++)
                {
                    var col = "col" + (j+1).ToString();
                    dt.Columns.Add(new DataColumn(col, vtdata[0, j].GetType()));
                }
            }
            var cols = dt.Columns;

            for (int i = start; i < n; i++)
            {
                var row = dt.NewRow();
                for (int j = 0; j < k; j++)
                {
                    row[cols[j].ColumnName] = vtdata[i, j];
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

    }
}
