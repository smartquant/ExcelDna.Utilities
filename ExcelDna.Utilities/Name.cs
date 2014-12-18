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

using ExcelDna.Integration;

namespace ExcelDna.Utilities
{

    public class Name
    {
        #region fields
        
        private Worksheet _worksheet;
        private Workbook _workbook;
        private string _name;
        
        #endregion

        #region constructors
        
        /// <summary>
        /// Name with worksheet scope
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="name"></param>
        internal Name(Worksheet worksheet, string name)
        {
            _worksheet = worksheet;
            _name = name;
        }
        
        /// <summary>
        /// Name with workbook scope
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        internal Name(Workbook workbook, string name)
        {
            _workbook = workbook;
            _name = name;
        }

        #endregion

        #region properties

        public string NameLocal
        {
            get { return _name; }
        }
        
        public string NameRef
        {
            get { return (_workbook != null) ? _workbook.Name + "!" + _name : _worksheet.SheetRef + "!" + _name; }
        }
        
        public bool IsLocalScope
        {
            get { return _worksheet != null; }
        }
        
        public bool IsGlobalScope
        {
            get { return _workbook != null; }
        }

        public string RefersTo
        {
            get
            {
                return (string)XlCall.Excel(XlCall.xlfGetName, this.NameRef, Type.Missing);
            }
        }

        #endregion

        #region static methods

        public static ExcelReference GetRange(string nameRef)
        {
            object result = XlCall.Excel(XlCall.xlfEvaluate, "=" + nameRef);

            return result as ExcelReference;
        }

        public static T GetValue<T>(string nameRef)
        {
            object result = XlCall.Excel(XlCall.xlfEvaluate, "=" + nameRef);

            ExcelReference r = result as ExcelReference;
            if (r != null)
                return r.GetValue().ConvertTo<T>();
            else
                return result.ConvertTo<T>();
        }

        public static T GetValue<T>(Workbook wb, string name)
        {
            return GetValue<T>(new Name(wb, name).RefersTo);
        }

        public static T GetValue<T>(Worksheet ws, string name)
        {
            return GetValue<T>(new Name(ws, name).RefersTo);
        }

        public static string NameRefersTo(string nameRef)
        {
            return (string)XlCall.Excel(XlCall.xlfGetName, nameRef, Type.Missing);
        }
        
        #endregion

        #region functions

        public override string ToString()
        {
            return _name;
        }

        #endregion
    }

}
