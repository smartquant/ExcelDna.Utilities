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
    public class Worksheet
    {
        #region fields
        
        private Workbook _workbook;
        private string _sheetname;
        
        #endregion

        #region constructors

        public Worksheet(Workbook wb, string sheet)
        {
            _workbook = wb;
            _sheetname = sheet;
        }

        public Worksheet(string sheetRef)
        {
            int pos = sheetRef.IndexOf(']');
            _workbook = new Workbook(sheetRef.Substring(1, pos - 1));
            _sheetname = sheetRef.Substring(pos + 1, sheetRef.Length - pos - 1);
        }

        #endregion

        #region properties

        public string Name
        {
            get { return _sheetname; }
            set
            {
                XlCall.Excel(XlCall.xlcWorkbookName, SheetRef, string.Format("[{0}]{1}", _workbook.Name, value));
                _sheetname = value;
            }
        }
        
        public Workbook Workbook
        {
            get { return _workbook; }
        }

        /// <summary>
        /// Full sheet reference as [workbook]SheetName
        /// </summary>
        public string SheetRef
        {
            get { return string.Format("[{0}]{1}", _workbook.Name, _sheetname); }
        }

        /// <summary>
        /// Returns all local names
        /// </summary>
        public Name[] NamesLocal
        {
            get
            {
                var sheetRef = this.SheetRef;

                //get all names (local and global) for this sheet
                object o = XlCall.Excel(XlCall.xlfNames, sheetRef);

                object[,] names = o as object[,];
                var list = new List<Name>();

                if (names != null)
                {
                    int n = names.GetLength(1);
                    for (int i = 0; i < n; i++)
                    {
                        string name = (string)names.GetValue(0, i);
                        string nameRef = string.Concat(sheetRef, "!", name);

                        //find out whether name is local or not
                        if ((bool)XlCall.Excel(XlCall.xlfGetName, nameRef, 2))
                            list.Add(new Name(this,name));

                    }
                }

                return list.ToArray();
            }
        }
        /// <summary>
        /// Returns all names for this worksheet (global and local)
        /// </summary>
        public Name[] Names
        {
            get
            {
                var sheetRef = this.SheetRef;

                //get all names (local and global) for this sheet
                object o = XlCall.Excel(XlCall.xlfNames, sheetRef);

                object[,] names = o as object[,];

                if (names != null)
                {
                    int n = names.GetLength(1);
                    Name[] res = new Name[n];
                    for (int i = 0; i < n; i++)
                    {
                        string name = (string)names.GetValue(0, i);
                        string nameRef = string.Concat(sheetRef, "!", name);

                        //find out whether name is local or not
                        if ((bool)XlCall.Excel(XlCall.xlfGetName, nameRef, 2))
                            res[i] = new Name(this, name);
                        else
                            res[i] = new Name(this.Workbook, name);
                    }
                    return res;
                }

                return new Name[0];
            }
        }
        #endregion

        #region static methods

        public static Worksheet ActiveSheet()
        {
            var res = XlCall.Excel(XlCall.xlfGetDocument, 76);
            return new Worksheet((string)res);
        }

        public static string ExtractSheetName(string sheetRef)
        {
            int pos = sheetRef.IndexOf(']');
            return sheetRef.Substring(pos + 1, sheetRef.Length - pos - 1);
        }

        public static string ExtractWorkbookName(string sheetRef)
        {
            int pos = sheetRef.IndexOf(']');
            return sheetRef.Substring(1, pos - 1);
        }

        #endregion

        #region functions

        public void Select()
        {
            this.Workbook.Activate();
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[,] { { this.SheetRef } }, Type.Missing, Type.Missing);
        }

        public void SelectAllCells()
        {
            XlCall.Excel(XlCall.xlcSelect, SheetRef + "!1:1048576", Type.Missing);
        }

        public ExcelReference Range(string range)
        {
            object result = XlCall.Excel(XlCall.xlfEvaluate, "='" + this.SheetRef + "'!" + range);
            return result as ExcelReference;
        }

        /// <summary>
        /// Defines the name on the ACTIVE sheet
        /// </summary>
        /// <param name="name"></param>
        /// <param name="refersto"></param>
        /// <param name="hidden"></param>
        /// <param name="local"></param>
        public Name AddName(string name, string refersto, bool hidden = false, bool local = true)
        {
            //Check whether this is the active sheet
            var ws = Worksheet.ActiveSheet();
            if (ws.Name != this.Name) return null;

            //DEFINE.NAME(name_text, refers_to, macro_type, shortcut_text, hidden, category, local)
            bool result = (bool)XlCall.Excel(XlCall.xlcDefineName, name, refersto, Type.Missing, Type.Missing, hidden, Type.Missing, local);
            if (result) return new Name(this, name);
            return null;
        }

        public void DeleteAllCells()
        {
            ExcelDna.Utilities.Name.GetRange(this.SheetRef + "!1:1048576").DeleteEntireRows();
        }

        #endregion
    }

}
