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
using System.Data;
using System.Linq;
using System.Text;
using System.IO;

using ExcelDna.Integration;



namespace ExcelDna.Utilities
{

    public static partial class ExcelReferenceEx
    {
        public static T GetValue<T>(this ExcelReference range)
        {
            return XLConversion.ConvertTo<T>(range.GetValue());
        }

        public static void DeleteEntireRows(this ExcelReference range)
        {
            //Shift_num	Result

            //1	Shifts cells left
            //2	Shifts cells up
            //3	Deletes entire row
            //4	Deletes entire column

            Action action = () =>
            {
                XlCall.Excel(XlCall.xlcEditDelete, 3);
            };

            XLApp.ActionOnSelectedRange(range, action);
        }

        public static void ClearContents(this ExcelReference range)
        {
            //int n = range.RowLast - range.RowFirst + 1, k = range.ColumnLast - range.ColumnFirst + 1;
            //var _empty = new object[n, k];
            //range.SetValue(_empty);

            Action action = () =>
            {
                XlCall.Excel(XlCall.xlcClear, 3);
            };

            XLApp.ActionOnSelectedRange(range, action);
        }

        /// <summary>
        /// Changes the color formatting of a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="backcolor">0: default color,number from 1 to 56 corresponding to the 56 area background colors in the Patterns tab of the Format Cells dialog box</param>
        /// <param name="forecolor">0: default color,number from 1 to 56 corresponding to the 56 area foreground colors in the Patterns tab of the Foramt Cells dialog box</param>
        /// <param name="pattern">0: auto pattern, pattern can be from 1 to 18</param>
        public static void FormatColor(this ExcelReference range, int backcolor = 0, int forecolor = 0, int pattern = 0)
        {
            //Should be called within screenupdating = false

            //change the pattern
            //PATTERNS(apattern, afore, aback, newui)
            Action action = () =>
            {
                XlCall.Excel(XlCall.xlcPatterns, pattern, forecolor, backcolor, true);
            };

            XLApp.ActionOnSelectedRange(range, action);
        }

        public static void FormatBorder(this ExcelReference range, xlBorderStyle outline = xlBorderStyle.NoBorder,
            xlBorderStyle left = xlBorderStyle.NoBorder,xlBorderStyle right = xlBorderStyle.NoBorder,
            xlBorderStyle top = xlBorderStyle.NoBorder, xlBorderStyle bottom = xlBorderStyle.NoBorder,
            int shade = 0, int outline_color = 0, int left_color = 0, int right_color = 0, int top_color = 0, int bottom_color = 0)
        {
            //BORDER(outline, left, right, top, bottom, shade, outline_color, left_color, right_color, top_color, bottom_color)

            Action action = () =>
            {
                XlCall.Excel(XlCall.xlcBorder, (int)left, (int)right, (int)top, (int)bottom, shade, outline_color, left_color, right_color, top_color, bottom_color);
            };

            XLApp.ActionOnSelectedRange(range, action);
        }

        public static void FormatNumber(this ExcelReference range, string format)
        {

            Action action = () =>
            {
                XlCall.Excel(XlCall.xlcFormatNumber, format);
            };

            XLApp.ActionOnSelectedRange(range, action);
        }

        public static string SheetRef(this ExcelReference range)
        {
            return (string)XlCall.Excel(XlCall.xlfGetCell, 62, range);
        }

        public static void Select(this ExcelReference range)
        {
            XlCall.Excel(XlCall.xlcFormulaGoto, range);
            XlCall.Excel(XlCall.xlcSelect, range, Type.Missing);
        }

        public static void Copy(this ExcelReference fromRange, ExcelReference toRange = null)
        {
            object to_range = (toRange == null) ? Type.Missing : toRange;
            XlCall.Excel(XlCall.xlcCopy, fromRange, to_range);
        }

        public static string RefersTo(this ExcelReference range)
        {
            object result = XlCall.Excel(XlCall.xlfGetCell, 6);
            return (string)result;
        }

        public static ExcelReference Resize(this ExcelReference range, int rows, int cols)
        {
            rows = (rows < 1) ? 1 : rows;
            cols = (cols < 1) ? 1 : cols;
            return new ExcelReference(range.RowFirst, range.RowFirst + rows-1, range.ColumnFirst, range.ColumnFirst + cols-1, range.SheetId);
        }

        public static ExcelReference Offset(this ExcelReference range, int rows, int cols)
        {
            return new ExcelReference(range.RowFirst + rows, range.RowLast + rows, range.ColumnFirst + cols, range.ColumnLast + cols, range.SheetId);
        }

        public static ExcelReference AddHeader(this ExcelReference range)
        {
            return new ExcelReference(range.RowFirst -1, range.RowLast, range.ColumnFirst, range.ColumnLast, range.SheetId);
        }

        public static List<T> ToList<T>(this ExcelReference range) where T : class
        {
            var items = new List<T>();
            XLObjectMapper.AddRange(items, range.GetValue());
            return items;
        }

        public static DataTable ToDataTable(this ExcelReference range, bool header = true)
        {
            return DataTableEx.CreateDataTable(range.GetValue(), header);
        }

        /// <summary>
        /// copy a variant matrix into an excel named range and adjust the size of the named range
        /// 1.) will copy the formatting of the first row when adding new rows
        /// 2.) will copy the formatting of the first row after the named range when removing rows
        /// </summary>
        /// <param name="vt">should be object[,] or simple data type</param>
        /// <param name="outRange"></param>
        /// <param name="localName">This is the local name of the output range (always ex header)</param>
        /// <param name="header">if there is a header the named range will start one cell below</param>
        /// <param name="ignoreFirstCell">will not fill the first cell; header will be inside the range if true</param>
        public static void Fill(this ExcelReference outRange, object vt, string localName = null, bool header = false, bool ignoreFirstCell = false)
        {
            var _vt = vt as object[,];
            if (_vt == null) _vt = new object[,] { { vt } };

            int name_offset = (header && ignoreFirstCell) ? 1 : 0;
            int origin_offset = ((header && !ignoreFirstCell) ? -1 : 0);
            int header_offset = (header) ? -1 : 0;
            int n = _vt.GetLength(0), k = _vt.GetLength(1);
            int m = outRange.RowLast - outRange.RowFirst + 1;

            //formatting
            bool localRange = !string.IsNullOrEmpty(localName);
            bool format = true;

            ExcelReference formatRange = null, newRange = null;

            if (m == 1 && localRange)
            {
                formatRange = Name.GetRange(outRange.SheetRef() + "!" + localName);
                if (formatRange == null)
                    format = false;
                else
                    m = formatRange.RowLast - formatRange.RowFirst + 1;
            }
            else if (m == 1)
                format = false;


            bool addRows = n + header_offset > m && format;
            bool removeRows = n + header_offset < m && format;


            int x0 = outRange.RowFirst + origin_offset, y0 = outRange.ColumnFirst; //output origin
            int x1 = outRange.RowFirst + name_offset, y1 = outRange.ColumnFirst; //name origin           

            bool updating = XLApp.ScreenUpdating;
            xlCalculation calcMode = XLApp.Calcuation;

            if (updating) XLApp.ScreenUpdating = false;

            try
            {
                var fillRange = new ExcelReference(x0, x0 + n - 1, y0, y0 + k - 1, outRange.SheetId);

                if (addRows)
                {
                    formatRange = new ExcelReference(x1, x1, y1, y1 + k - 1, outRange.SheetId); //first row
                    newRange = new ExcelReference(x1, x1 + n + header_offset - 1, y1, y1 + k - 1, outRange.SheetId);
                }
                if (removeRows)
                {
                    formatRange = new ExcelReference(x1 + m, x1 + m, y1, y1 + k - 1, outRange.SheetId); //last row + 1
                    newRange = new ExcelReference(x1 + n + header_offset, x1 + m - 1, y1, y1 + k - 1, outRange.SheetId);
                    newRange.ClearContents();
                }

                //set the range except the first cell
                if (ignoreFirstCell && n > 1)
                {
                    //first row
                    if (k > 1)
                    {
                        object[,] first = new object[1, k - 1];
                        for (int i = 0; i < k - 1; i++)
                            first[0, i] = _vt[0, i + 1];
                        fillRange.Offset(0, 1).Resize(1, k - 1).SetValue(first);
                    }
                    //all other rows
                    object[,] rest = new object[n - 1, k];
                    for (int i = 1; i < n; i++)
                        for (int j = 0; j < k; j++)
                            rest[i - 1, j] = _vt[i, j];
                    fillRange.Offset(1, 0).Resize(n - 1, k).SetValue(rest);
                }
                else if (!ignoreFirstCell)
                    fillRange.SetValue(_vt);


                //set name
                if (localRange)
                {
                    Action action = () =>
                    {
                        string sheetref = (string)XlCall.Excel(XlCall.xlSheetNm, outRange);
                        Worksheet sheet = new Worksheet(sheetref);

                        //re-color
                        if (addRows || removeRows)
                        {
                            formatRange.Select();
                            XLApp.Copy();
                            newRange.Select();
                            XLApp.PasteSpecial(xlPasteType.PasteFormats);
                        }

                        string reference = string.Format("={4}!R{0}C{2}:R{1}C{3}", x1 + 1, x1 + n + header_offset, y1 + 1, y1 + k, sheetref);

                        //DEFINE.NAME(name_text, refers_to, macro_type, shortcut_text, hidden, category, local)
                        XlCall.Excel(XlCall.xlcDefineName, sheet.Name + "!" + localName, reference, Type.Missing, Type.Missing, false, Type.Missing, true);
                    };
                    XLApp.ActionOnSelectedRange(fillRange, action);
                }
            }
            finally
            {
                if (updating) XLApp.ScreenUpdating = true;
            }

        }

        //TODO: these functions should be able to paste chunks of data say 5000 lines per chunk rather than converting everything in one array

        public static void Fill(this ExcelReference outRange, DataTable dt, string localName = null, bool header = false, bool ignoreFirstCell = false)
        {
            var vt = dt.ToVariant(header);
            Fill(outRange, vt, localName, header, ignoreFirstCell);
        }

        public static void Fill<T>(this ExcelReference outRange, IEnumerable<T> items, string localName = null, bool header = false, bool ignoreFirstCell = false) where T : class
        {
            var vt = items.ToVariant(header);
            Fill(outRange, vt, localName, header, ignoreFirstCell);
        }
    }
}
