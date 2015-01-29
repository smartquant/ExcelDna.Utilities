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
using System.Linq;
using System.Text;
using System.IO;

using ExcelDna.Integration;



namespace ExcelDna.Utilities
{

    public static partial class XLApp
    {
        private static bool _screenupdating = true;

        #region properties

        public static Workbook[] Workbooks
        {
            get
            {
                var list = new List<Workbook>();

                object o = XlCall.Excel(XlCall.xlfDocuments);
                object[,] docs = o as object[,];

                if (docs != null)
                    for (int i = 0; i < docs.GetLength(1); i++)
                        list.Add(new Workbook((string)docs.GetValue(0, i)));

                return list.ToArray();
            }
        }

        public static string DefaultDateFormat
        {
            get
            {
                var result = XlCall.Excel(XlCall.xlfGetWorkspace, 37) as object[,];

                int i = 16;
                string date_seperator = (string)result[0, i++];
                string time_seperator = (string)result[0, i++];
                string year_symbol = (string)result[0, i++];
                string month_symbol = (string)result[0, i++];
                string day_symbol = (string)result[0, i++];
                string hour_symbol = (string)result[0, i++];
                string minute_symbol = (string)result[0, i++];
                string second_symbol = (string)result[0, i++];
                //32	Number indicating the date order
                //0 = Month-Day-Year
                //1 = Day-Month-Year
                //2 = Year-Month-Day
                double date_order = (double)result[0, 31];

                day_symbol = day_symbol + day_symbol;
                month_symbol = month_symbol + month_symbol;
                year_symbol = string.Concat(year_symbol, year_symbol, year_symbol, year_symbol);

                if (date_order == 0)
                    return month_symbol + date_seperator + day_symbol + date_seperator + year_symbol;
                else if (date_order == 1)
                    return day_symbol + date_seperator + month_symbol + date_seperator + year_symbol;
                else
                    return year_symbol + date_seperator + month_symbol + date_seperator + day_symbol;
            }

        }

        #endregion

        #region Message bar
        /// <summary>
        /// Similar to excel pass an empty string to reset message bar
        /// </summary>
        /// <param name="message"></param>
        public static void MessageBar(string message)
        {
            bool display = !string.IsNullOrEmpty(message);
            XlCall.Excel(XlCall.xlcMessage, display, message);
        }
        
        public static void MessageBar(string message, params object[] obj)
        {
            MessageBar(string.Format(message,obj));
        }

        #endregion

        #region calculation

        public static bool ScreenUpdating
        {
            set
            {
                XlCall.Excel(XlCall.xlcEcho, value);
                _screenupdating = value;
            }
            get
            {
                //return (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 40);;
                return _screenupdating;
            }
        }

        public static xlCalculation Calcuation
        {
            get 
            {
                var result = XlCall.Excel(XlCall.xlfGetDocument, 14);
                return (xlCalculation)(int)(double)result;
            }
            set
            {
                //get all calculation settings for the function call OPTIONS.CALCULATION
                object[,] result = XlCall.Excel(XlCall.xlfGetDocument, new object[,]{{14,15,16,17,18,19,20,33,43}}) as object[,];
                object[] pars = result.ToVector<object>();
                pars[0]=(int)value;
                var retval = XlCall.Excel(XlCall.xlcOptionsCalculation, pars);
            }
        }
        
        public static void CalculateNow()
        {
            XlCall.Excel(XlCall.xlcCalculateNow);
        }
        
        public static void CalculateDocument()
        {
            XlCall.Excel(XlCall.xlcCalculateDocument);
        }

        #endregion

        #region wrapping actions

        /// <summary>
        /// Wrapper for macro functions that need a selection; remembers old selection
        /// </summary>
        /// <param name="range"></param>
        /// <param name="action"></param>
        public static void ActionOnSelectedRange(this ExcelReference range, Action action)
        {
            bool updating = ScreenUpdating;

            try
            {
                if (updating) ScreenUpdating = false;

                //remember the current active cell 
                object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);

                //select caller range AND workbook
                string rangeSheet = (string)XlCall.Excel(XlCall.xlSheetNm, range);

                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { rangeSheet });
                XlCall.Excel(XlCall.xlcSelect, range);

                action.Invoke();

                //go back to old selection
                XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
            }
            finally
            {
                if (updating) XLApp.ScreenUpdating = true;
            }
        }

        public static void ReturnToSelection(Action action)
        {
            bool updating = ScreenUpdating;

            try
            {
                if (updating) ScreenUpdating = false;

                //remember the current active cell 
                object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                
                action.Invoke();

                //go back to old selection
                XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
            }
            finally
            {
                if (updating) XLApp.ScreenUpdating = true;
            }
        }

        public static void NoCalcAndUpdating(this Action action)
        {
            xlCalculation calcMode = Calcuation;
            bool updating = ScreenUpdating;           

            try
            {
                if (updating) ScreenUpdating = false;
                if (calcMode != xlCalculation.Manual) Calcuation = xlCalculation.Manual;

                action.Invoke();
            }
            finally
            {
                if (updating) XLApp.ScreenUpdating = true;
                if (calcMode != xlCalculation.Manual) XLApp.Calcuation = calcMode;
            }
        }

        #endregion

        #region copy paste

        public static void PasteSpecial(xlPasteType type = xlPasteType.PasteAll, xlPasteAction action = xlPasteAction.None, bool skip_blanks = false, bool transpose=false)
        {
            XlCall.Excel( XlCall.xlcPasteSpecial,(int)type,(int)action,skip_blanks,transpose);
        }

        public static void Copy()
        {
            XlCall.Excel(XlCall.xlcCopy, Type.Missing, Type.Missing);
        }

        #endregion

        public static void SelectRange(string rangeRef)
        {
            XlCall.Excel(XlCall.xlcSelect, rangeRef, Type.Missing);
        }


    }
}
