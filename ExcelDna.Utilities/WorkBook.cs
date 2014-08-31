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
using System.IO;

using ExcelDna.Integration;

namespace ExcelDna.Utilities
{
    public class Workbook
    {
        #region fields

        private string _path;
        private string _workbook;

        #endregion

        #region constructor

        public Workbook(string filename)
        {
            Init(filename);
        }

        private void Init(string filename)
        {
            FileInfo finfo = new FileInfo(filename);
            _path = finfo.DirectoryName + System.IO.Path.DirectorySeparatorChar;
            _workbook = finfo.Name;
        }

        #endregion

        #region static methods

        public static Workbook CreateNew()
        {
            XlCall.Excel(XlCall.xlcNew, 5);
            var res = XlCall.Excel(XlCall.xlfGetDocument, 88);
            return new Workbook((string)res);
        }

        public static Workbook ActiveWorkbook()
        {
            var res = XlCall.Excel(XlCall.xlfGetDocument, 88);
            return new Workbook((string)res);
        }

        public static Workbook Open(string path, xlUpdateLinks update_links = xlUpdateLinks.Never, bool read_only = false, string password = null)
        {
            XlCall.Excel(XlCall.xlcOpen, path, (int)update_links, read_only, Type.Missing, password, Type.Missing, true, 2,
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Workbook wb = new Workbook(path);
            return wb;
        }

        #endregion

        #region properties

        public string Path
        {
            get { return _path; }
        }

        public string Name
        {
            get { return _workbook; }
        }

        public bool IsOpen
        {
            get
            {
                var res = XlCall.Excel(XlCall.xlfGetDocument, 3, _workbook);
                Type t = res.GetType();
                return (t != typeof(ExcelError));
            }
        }

        public bool IsReadonly
        {
            get
            {
                var res = XlCall.Excel(XlCall.xlfGetDocument, 5, _workbook);
                Type t = res.GetType();
                if (t == typeof(ExcelError)) return true;
                return (bool)res;
            }
        }

        #endregion 

        #region close save

        public void Activate()
        {
            XlCall.Excel(XlCall.xlcActivate, this.Name, Type.Missing);
        }

        public void Close(bool saveChanges = true, bool routeFile = false)
        {
            Activate();
            XlCall.Excel(XlCall.xlcClose, saveChanges, routeFile);
        }

        public string GetPath()
        {
            _path = (string)XlCall.Excel(XlCall.xlfGetDocument, _workbook, 2);
            return _path;
        }

        public void SaveAs(string path, string password = null, string write_password = null, bool read_only = false)
        {
            object pwd = string.IsNullOrEmpty(password) ? Type.Missing : password;
            object write_pwd = string.IsNullOrEmpty(write_password) ? Type.Missing : write_password;

            XlCall.Excel(XlCall.xlcSaveAs, path, Type.Missing, pwd, Type.Missing, write_pwd, read_only);
            Init(path);
        }

        public void Save()
        {
            Activate();
            XlCall.Excel(XlCall.xlcSave);
        }

        #endregion

        #region functions

        public Worksheet AddWorksheet()
        {
            Activate();
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);
            return Worksheet.ActiveSheet();
        }

        public Worksheet[] Worksheets
        {
            get
            {
                object[,] sheetnames =
                    (object[,])XlCall.Excel(XlCall.xlfGetWorkbook, 1, this.Name);
                int n = sheetnames.GetLength(1);
                Worksheet[] sheets = new Worksheet[n];

                for (int j = 0; j < n; j++)
                {
                    sheets[j] = new Worksheet(sheetnames[0, j].ToString());
                }
                return sheets;
            }
        }
        
        public String[] SheetRefs
        {
            get
            {
                object[,] sheetnames =
                    (object[,])XlCall.Excel(XlCall.xlfGetWorkbook, 1, this.Name);
                int n = sheetnames.GetLength(1);
                string[] sheets = new string[n];

                for (int j = 0; j < n; j++)
                {
                    sheets[j] = sheetnames[0, j].ToString();
                }
                return sheets;
            }
        }

        public String[] SheetNames
        {
            get
            {
                var sheetrefs = this.SheetRefs;
                int n = sheetrefs.Length;
                string[] result = new string[n];

                for (int i = 0; i < n; i++)
                    result[i] = Worksheet.ExtractSheetName(sheetrefs[i]);

                return result;
            }
        }

        /// <summary>
        /// References a sheet from this workbook, however DOES NOT TEST wheter sheet actually exists for
        /// performance reasons
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Worksheet this[string sheetName]
        {
            get
            {
                return new Worksheet(this, sheetName);
            }
        }

        #endregion
    }
}
