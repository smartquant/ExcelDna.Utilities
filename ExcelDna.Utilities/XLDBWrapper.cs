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
using System.Data.Common;

using ExcelDna.Integration;


namespace ExcelDna.Utilities
{

    public static class XLDBWrapper
    {


        public delegate object FieldFormatter(object obj);

        public static object DefaultFieldFormatter(object obj)
        {
            return (obj is Array) ? "Array data" : obj;
        }
        /// <summary>
        /// Transforms a DBDataReader object into an Excel Variant datatype
        /// Note: Returned number of rows should not be too large (say below 10000)
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="properties">Field names to be selected; null for taking alll default</param>
        /// <param name="colheaders">The columnnames in the variant output; null for taking values in properts</param>
        /// <param name="header">Should column names be written to the variant output</param>
        /// <param name="rows">the number of rows (ex header)</param>
        /// <param name="cols">the number of columns of the variant</param>
        /// <returns></returns>
        public static object[,] ToVariant(this DbDataReader reader, string[] properties=null, string[] colheaders=null, FieldFormatter formatter=null, bool header=true)
        {
            string[] props;
            string[] colheader;

            if (properties == null)
            {
                List<string> sc = new List<string>();
                for (int i = 0; i < reader.FieldCount; i++)
                    sc.Add(reader.GetName(i));
                props = sc.ToArray();

            }
            else
            {
                props = properties;
            }

            if (header && (colheaders == null || colheaders.Length != props.Length))
                colheader = props;
            else
                colheader = colheaders;
            if (formatter == null)
                formatter = DefaultFieldFormatter;

            int iheader = (header) ? 1 : 0;
            int @base = 0;

            int[] idx = new int[props.Length];

            for (int i = 0; i < props.Length; i++)
                idx[i] = reader.GetOrdinal(props[i]);

            List<object[]> oList = new List<object[]>();
            int nFields = props.Length;
            while (reader.Read())
            {
                object[] vals = new object[nFields];
                for (int i = 0; i < nFields; i++)
                    vals[i] = reader.GetValue(idx[i]);
                oList.Add(vals);
            }

            int n = oList.Count + iheader, m = nFields;
            int[] lbound = { @base, @base };
            int[] lengths = { n, m };

            object[,] @out = (object[,])Array.CreateInstance(typeof(object), lengths, lbound);

            int j = @base;
            if (header)
                foreach (string col in colheader)
                    @out[@base, j++] = col;


            int k = @base + iheader;
            for (int i = 0; i < oList.Count; i++)
                for (int l = 0; l < m; l++)
                    @out[i + @base + iheader, l + @base] = formatter(oList[i][l]);
            //rows = n; cols = m;
            return @out;

        }



        /// <summary>
        /// Directly load the the rows from a dbreader into a list of xlobjects
        /// Important: The order of the returned columns must match the one of T column mapping
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="reader"></param>
        /// <returns></returns>
        public static List<T> ToXLObjectList<T>(this DbDataReader reader) where T : class
        {
            List<T> list = new List<T>();
            Type t = typeof(T);
            var map = XLObjectMapper.GetObjectMapping<T>();

            while (reader.Read())
            {
                T instance = (t.GetConstructor(Type.EmptyTypes) != null) ? (T)Activator.CreateInstance(t, new object[0])
                            : Activator.CreateInstance<T>();
                int nFields = reader.FieldCount;

                for (int i = 0; i < nFields; i++)
                    map.SetColumn(instance, i, reader.GetValue(i));
                list.Add(instance);
            }
            return list;
        }


    }
}
