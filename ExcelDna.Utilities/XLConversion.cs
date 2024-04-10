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

using System.Collections.Concurrent;

using ExcelDna.Integration;

namespace ExcelDna.Utilities;

public static class XLConversion
{
    #region object conversion
    //a thread-safe way to hold default instances created at run-time
    private static ConcurrentDictionary<Type, object> typeDefaults = new();

    private static object GetDefault(Type type)
    {
        return type.IsValueType ? typeDefaults.GetOrAdd(type, t => Activator.CreateInstance(t)) : null;
    }

    public static object ConvertTo(this object vt, Type toType)
    {
        if (vt == null) return GetDefault(toType);
        Type fromType = vt.GetType();
        if (fromType == typeof(DBNull)) return GetDefault(toType);

        if (fromType == typeof(ExcelDna.Integration.ExcelEmpty) || fromType == typeof(ExcelDna.Integration.ExcelError) || fromType == typeof(ExcelDna.Integration.ExcelMissing))
            return GetDefault(toType);

        if (fromType == typeof(ExcelReference))
        {
            ExcelReference r = (ExcelReference)vt;
            object val = r.GetValue();
            return ConvertTo(val, toType);
        }

        //acount for nullable types
        toType = Nullable.GetUnderlyingType(toType) ?? toType;

        if (toType == typeof(DateTime))
        {
            DateTime dt = DateTime.FromOADate(0.0);
            if (fromType == typeof(DateTime))
                dt = (DateTime)vt;
            else if (fromType == typeof(double))
                dt = DateTime.FromOADate((double)vt);
            else if (fromType == typeof(string))
            {
                DateTime result;
                if (DateTime.TryParse((string)vt, out result))
                    dt = result;
            }
            return Convert.ChangeType(dt, toType);
        }
        else if (toType == typeof(XLDate))
        {
            XLDate dt = 0.0;
            if (fromType == typeof(DateTime))
                dt = (DateTime)vt;
            else if (fromType == typeof(double))
                dt = (double)vt;
            else if (fromType == typeof(string))
            {
                DateTime result;
                if (DateTime.TryParse((string)vt, out result))
                    dt = result;
                else
                    dt = 0.0;
            }
            else
                dt = (double)Convert.ChangeType(vt, typeof(double));
            return Convert.ChangeType(dt, toType);
        }
        else if (toType == typeof(double))
        {
            double dt = 0.0;
            if (fromType == typeof(double))
                dt = (double)vt;
            else if (fromType == typeof(DateTime))
                dt = ((DateTime)vt).ToOADate();
            else if (fromType == typeof(string))
                double.TryParse((string)vt, out dt);
            else
                dt = (double)Convert.ChangeType(vt, typeof(double));
            return Convert.ChangeType(dt, toType);
        }
        else if (toType.IsEnum)
        {
            try
            {
                return Enum.Parse(toType, vt.ToString(), true);
            }
            catch (Exception)
            {
                return GetDefault(toType);
            }

        }
        else
            return Convert.ChangeType(vt, toType);


    }

    public static T ConvertTo<T>(this object vt)
    {
        if (vt == null) return default(T);

        Type toType = typeof(T);
        Type fromType = vt.GetType();
        if (fromType == typeof(DBNull)) return default(T);

        if (fromType == typeof(ExcelDna.Integration.ExcelEmpty) || fromType == typeof(ExcelDna.Integration.ExcelError) || fromType == typeof(ExcelDna.Integration.ExcelMissing))
            return default(T);

        if (fromType == typeof(ExcelReference))
        {
            ExcelReference r = (ExcelReference)vt;
            object val = r.GetValue();
            return ConvertTo<T>(val);
        }

        //acount for nullable types
        toType = Nullable.GetUnderlyingType(toType) ?? toType;

        if (toType == typeof(DateTime))
        {
            DateTime dt = DateTime.FromOADate(0.0);
            if (fromType == typeof(DateTime))
                dt = (DateTime)vt;
            else if (fromType == typeof(double))
                dt = DateTime.FromOADate((double)vt);
            else if (fromType == typeof(string))
            {
                DateTime result;
                if (DateTime.TryParse((string)vt, out result))
                    dt = result;
            }
            //note this will work also if T is nullable
            return (T)Convert.ChangeType(dt, toType);
        }
        else if (toType == typeof(XLDate))
        {
            XLDate dt = 0.0;
            if (fromType == typeof(DateTime))
                dt = (DateTime)vt;
            else if (fromType == typeof(double))
                dt = (double)vt;
            else if (fromType == typeof(string))
            {
                DateTime result;
                if (DateTime.TryParse((string)vt, out result))
                    dt = result;
                else
                    dt = 0.0;
            }
            else
                dt = (double)Convert.ChangeType(vt, typeof(double));
            return (T)Convert.ChangeType(dt, toType);
        }
        else if (toType == typeof(double))
        {
            double dt = 0.0;
            if (fromType == typeof(double))
                dt = (double)vt;
            else if (fromType == typeof(DateTime))
                dt = ((DateTime)vt).ToOADate();
            else if (fromType == typeof(string))
                double.TryParse((string)vt, out dt);
            else
                dt = (double)Convert.ChangeType(vt, typeof(double));
            return (T)Convert.ChangeType(dt, toType);
        }
        else if (toType.IsEnum)
        {
            try
            {
                return (T)Enum.Parse(typeof(T), vt.ToString(), true);
            }
            catch (Exception)
            {
                return default(T);
            }

        }
        else
            return (T)Convert.ChangeType(vt, toType);

    }

    public static void ConvertVT<T>(this object vt, out T value)
    {
        value = vt.ConvertTo<T>();
    }

    public static T[] ToVector<T>(this object vt)
    {
        if (vt is Array) return ToVector<T>(vt as object[,]);

        T[] retval = new T[1];
        vt.ConvertVT(out retval[0]);

        return retval;
    }

    public static T[] ToVector<T>(this object[,] vt)
    {
        int n = vt.GetLength(0), k = vt.GetLength(1);
        int l = 0;

        T[] @out = new T[n * k];

        for (int i = 0; i < n; i++)
            for (int j = 0; j < k; j++)
            {
                vt.GetValue(i, j).ConvertVT<T>(out @out[l]);
                l++;
            }

        return @out;
    }

    public static T[,] ToMatrix<T>(this object vt)
    {
        if (vt is Array) return ToMatrix<T>(vt as object[,]);

        T[,] retval = new T[1, 1];
        vt.ConvertVT(out retval[0, 0]);

        return retval;
    }

    public static T[,] ToMatrix<T>(this object[,] vt)
    {
        int n = vt.GetLength(0), k = vt.GetLength(1);

        T[,] @out = new T[n, k];

        for (int i = 0; i < n; i++)
            for (int j = 0; j < k; j++)
            {
                vt.GetValue(i, j).ConvertVT<T>(out @out[i, j]);
            }

        return @out;
    }

    public static object ToVariant<T>(this T[,] vt)
    {
        int n = vt.GetLength(0), k = vt.GetLength(1);
        object[,] @out = new object[n, k];

        for (int i = 0; i < n; i++)
            for (int j = 0; j < k; j++)
            {
                @out[i, j] = vt.GetValue(i, j);
            }

        return @out;
    }

    #endregion
}

