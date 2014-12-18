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
using System.Collections.Concurrent;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ExcelDna.Utilities
{

    /// <summary>
    /// Interface that allows to enumerate and name the fields of an object
    /// typically used for interacting with excel ranges and Lists of strongly typed objects
    /// that cannot be enumerated easily or automatically through reflection
    /// XLObjectMapping will use these values instead of reflection if the type implements this interface
    /// </summary>
    public interface IXLObjectMapping
    {
        /// <summary>
        /// Number of columns
        /// </summary>
        /// <returns></returns>
        int ColumnCount();
        /// <summary>
        /// Column name for index 
        /// </summary>
        /// <param name="index">zero based</param>
        /// <returns></returns>
        string ColumnName(int index);
        /// <summary>
        /// Indexed getter
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        object GetColumn(int index);
        /// <summary>
        /// Indexed setter
        /// </summary>
        /// <param name="index"></param>
        /// <param name="RHS"></param>
        void SetColumn(int index, object RHS);
    }


    /// <summary>
    /// Class to store the column mappings for poco objects so we can map excel ranges to objects of type T
    /// </summary>
    public static class XLObjectMapper
    {
        #region mappings

        private static ConcurrentDictionary<Type, XLObjectMapping> _types = new ConcurrentDictionary<Type, XLObjectMapping>();

        public static ConcurrentDictionary<Type, XLObjectMapping> ObjectMappings
        {
            get { return _types; }
        }

        public static XLObjectMapping GetObjectMapping<T>()
        {
            var t = typeof(T);

            return _types.GetOrAdd(t, f => new XLObjectMapping(t));
        }

        public static void SetObjectMapping(XLObjectMapping mapping)
        {
            _types.AddOrUpdate(mapping.MappedType, mapping, (t, m) => mapping);
        }

        #endregion

        #region utility functions

        public static object[,] ToVariant<T>(this IEnumerable<T> items, bool header = false) where T : class
        {
            if (items.Count() == 0) return new object[1, 1] { { 0 } };

            T obj = items.First();
            var map = GetObjectMapping<T>();
            int n = items.Count(), cols = map.Columns;
            int rows = header ? n + 1 : n;
            int start = header ? 1 : 0;


            object[,] retval = new object[rows, cols];

            if (header)
            {
                for (int i = 0; i < cols; i++)
                    retval[0, i] = map.Colnames[i];
            }

            foreach (T item in items)
            {
                for (int j = 0; j < cols; j++)
                    retval[start, j] = map.GetColumn(item,j);
                start++;
            }

            return retval;

        }

        public static String ToString<T>(this IEnumerable<T> items, string sep = ",", bool header = false) where T : class
        {
            var sb = new StringBuilder();
            if (items.Count() == 0) return string.Empty;

            T obj = items.First();
            var map = GetObjectMapping<T>();
            int n = items.Count(), cols = map.Columns;
            int rows = header ? n + 1 : n;
            int start = header ? 1 : 0;


            if (header)
            {
                for (int i = 0; i < cols; i++)
                    sb.Append(map.Colnames[i] + sep);
                sb.Length -= sep.Length;
                sb.AppendLine();
            }

            foreach(var item in items)
            {
                for (int j = 0; j < cols; j++)
                    sb.Append(map.GetColumn(item,j) + sep);

                sb.Length -= sep.Length;
                sb.AppendLine();
                start++;
            }
            return sb.ToString();
        }

        public static void AddRange<T>(this ICollection<T> items, object vt) where T : class
        {
            object[,] vtdata = vt as object[,];

            int n = vtdata.GetLength(0);
            int k = vtdata.GetLength(1);

            Type t = typeof(T);
            var map = GetObjectMapping<T>();

            for (int i = 0; i < n; i++)
            {
                T instance = (t.GetConstructor(Type.EmptyTypes) != null) ? (T)Activator.CreateInstance(t,new object[0]) 
                            : Activator.CreateInstance<T>();

                for (int j = 0; j < k; j++)
                {
                    map.SetColumn(instance, j, vtdata[i, j]);
                }
                items.Add(instance);
            }
        }

        /// <summary>
        /// Reads the a delimeted string into the fields of a IXLRow object
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="s"></param>
        /// <param name="sep"></param>
        /// <returns></returns>
        public static T DeserializeFromString<T>(this string s, string sep = "|") where T : class
        {
            Type t = typeof(T);
            T instance = (t.GetConstructor(Type.EmptyTypes) != null) ? (T)Activator.CreateInstance(t, new object[0])
                        : Activator.CreateInstance<T>();
            var map = GetObjectMapping<T>();
            string[] vt = s.Split(new string[] { sep }, StringSplitOptions.None);
            for (int i = 0; i < vt.Length; i++)
            {
                if (!string.IsNullOrEmpty(vt[i]))
                    map.SetColumn(instance, i, vt[i]);
            }
            return instance;
        }
        /// <summary>
        /// Puts the contents of the IXLRow object's fields into a single line string that can be typically written into a Excel Name
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="sep"></param>
        /// <returns></returns>
        public static string SerializeToString<T>(this T obj, string sep = "|") where T : class
        {
            StringBuilder sb = new StringBuilder();
            var map = GetObjectMapping<T>();

            for (int i = 0; i < map.Columns; i++)
            {
                if (map.GetColumn(obj,i) != null)
                    sb.Append(map.GetColumn(obj,i).ToString() + sep);
                else
                    sb.Append(sep);
            }
            return sb.ToString(0, sb.Length - sep.Length);
        }

        public static string GetHeader<T>(this T obj, string sep = ",") where T : class
        {
            var map = GetObjectMapping<T>();
            return string.Join(sep, map.Colnames);
        }


        #endregion

    }

}
