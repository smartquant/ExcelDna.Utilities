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
        /// <param name="vt">must be object[,]</param>
        /// <param name="outRange"></param>
        /// <param name="localName">This is the local name</param>
        /// <param name="header"></param>
        public static void Fill(this ExcelReference outRange, object vt, string localName, bool header = false)
        {
            var _vt = vt as object[,];
            if (_vt == null) _vt = new object[,] { { vt } };

            int n = _vt.GetLength(0), k = _vt.GetLength(1);
            int m = outRange.RowLast - outRange.RowFirst + 1;

            bool addRows = n > m;
            bool removeRows = n < m;

            bool updating = XLApp.ScreenUpdating;
            if (updating) XLApp.ScreenUpdating = false;

            

            int start = (header) ? -1 : 0;

            ExcelReference formatRange = null, newRange = null;

            var _outrange = new ExcelReference(outRange.RowFirst + start, outRange.RowFirst + start + n - 1,
                outRange.ColumnFirst, outRange.ColumnFirst + k - 1, outRange.SheetId);

            if (addRows)
            {
                formatRange = new ExcelReference(outRange.RowFirst + start, outRange.RowFirst + start,
                    outRange.ColumnFirst, outRange.ColumnFirst + k - 1, outRange.SheetId);
                newRange = _outrange;
            }
            if (removeRows)
            {
                formatRange = new ExcelReference(outRange.RowFirst + start + m, outRange.RowFirst + start + m,
                    outRange.ColumnFirst, outRange.ColumnFirst + k - 1, outRange.SheetId);
                newRange = new ExcelReference(outRange.RowFirst + start + n, outRange.RowFirst + start + m - 1,
                    outRange.ColumnFirst, outRange.ColumnFirst + k - 1, outRange.SheetId);
                newRange.ClearContents();
            }

            _outrange.SetValue(_vt);

            //set name
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

                string reference = string.Format("={4}!R{0}C{2}:R{1}C{3}", outRange.RowFirst + 1,
                    outRange.RowFirst + start + n, outRange.ColumnFirst + 1, outRange.ColumnFirst + k, sheetref);

                //DEFINE.NAME(name_text, refers_to, macro_type, shortcut_text, hidden, category, local)
                XlCall.Excel(XlCall.xlcDefineName, sheet.Name + "!" + localName, reference, Type.Missing, Type.Missing, false, Type.Missing, true);
            };
            XLApp.ActionOnSelectedRange(_outrange, action);

            if (updating) XLApp.ScreenUpdating = true;

        }

        //TODO: these function should be able to paste chunks of data say 5000 lines per chunk rather than converting everything in one array

        public static void Fill(this ExcelReference outRange, DataTable dt, string localName, bool header = false)
        {
            var vt = dt.ToVariant(header);
            Fill(outRange, vt, localName, header);
        }

        public static void Fill<T>(this ExcelReference outRange, IEnumerable<T> items, string localName, bool header = false) where T: class
        {
            var vt = items.ToVariant(header);
            Fill(outRange, vt, localName, header);
        }
    }
}
