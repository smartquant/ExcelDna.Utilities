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


namespace ExcelDna.Utilities
{
    /// <summary>
    /// Convenience struct to work directly with Excel double dates rather than using DateTime.FromOADate() conversions
    /// for certain calculations there exist a direct and fast operation without going through all the conversions
    /// Implicit operators make sure that class is equivalent to DateTime and double.
    /// </summary>
    public struct XLDate : IComparable
    {

        private double _xlDate;

        #region constructors

        public XLDate(double xlDate)
        {
            _xlDate = xlDate;
        }
        public XLDate(XLDate xlDate)
        {
            _xlDate = xlDate._xlDate;
        }
        public XLDate(DateTime dateTime)
        {
            _xlDate = dateTime.ToOADate();
        }
        public XLDate(int year, int month, int day, int hour=0, int minute=0, int second=0, int millisecond=0)
        {
            _xlDate = new DateTime(year, month, day, hour, minute, second, millisecond).ToOADate();
        }

        #endregion

        #region properties

        public XLDate Date
        {
            get { return Math.Floor(this._xlDate); }
        }

        public int Year
        {
            get { return DateTime.FromOADate(_xlDate).Year; }
        }
        public int Month
        {
            get { return DateTime.FromOADate(_xlDate).Month; }
        }
        public int Day
        {
            get { return DateTime.FromOADate(_xlDate).Day; }
        }
        public DayOfWeek DayOfWeek
        {
            get { return DateTime.FromOADate(_xlDate).DayOfWeek; }
        }
        public int DayOfYear
        {
            get { return DateTime.FromOADate(_xlDate).DayOfYear; }
        }
        public int Hour
        {
            get { return DateTime.FromOADate(_xlDate).Hour; }
        }
        public int Minute
        {
            get { return DateTime.FromOADate(_xlDate).Minute; }
        }
        public int Second
        {
            get { return DateTime.FromOADate(_xlDate).Second; }
        }
        public int Millisecond
        {
            get { return DateTime.FromOADate(_xlDate).Millisecond; }
        }

        #endregion

        #region Date math

        public XLDate AddMilliseconds(double value)
        {
            return new XLDate(_xlDate + value / 86400000.0);
        }

        public XLDate AddSeconds(double value)
        {
            return new XLDate(_xlDate + value / 86400.0);
        }

        public XLDate AddMinutes(double value)
        {
            return new XLDate(_xlDate + value / 1440.0);
        }

        public XLDate AddHours(double value)
        {
            return new XLDate(_xlDate + value / 24.0);
        }

        public XLDate AddDays(double value)
        {
            return new XLDate(_xlDate + value); 
        }

        public XLDate AddMonths(int value)
        {
            return new XLDate(DateTime.FromOADate(_xlDate).AddMonths(value));
        }

        public XLDate AddYears(int value)
        {
            return new XLDate(DateTime.FromOADate(_xlDate).AddYears(value));
        }

        #endregion

        #region Operators

        public static double operator -(XLDate lhs, XLDate rhs)
        {
            return lhs._xlDate - rhs._xlDate;
        }

        public static XLDate operator -(XLDate lhs, double rhs)
        {
            lhs._xlDate -= rhs;
            return lhs;
        }

        public static XLDate operator +(XLDate lhs, double rhs)
        {
            lhs._xlDate += rhs;
            return lhs;
        }
        
        public static XLDate operator +(XLDate d, TimeSpan t)
        {
            XLDate date = new XLDate(d);
            d.AddMilliseconds(t.TotalMilliseconds);
            return date;
        }

        public static XLDate operator -(XLDate d, TimeSpan t)
        {
            XLDate date = new XLDate(d);
            d.AddMilliseconds(-t.TotalMilliseconds);
            return date;
        }
        
        public static XLDate operator ++(XLDate xDate)
        {
            xDate._xlDate += 1.0;
            return xDate;
        }

        public static XLDate operator --(XLDate xDate)
        {
            xDate._xlDate -= 1.0;
            return xDate;
        }

        public static implicit operator double(XLDate xDate)
        {
            return xDate._xlDate;
        }
        
        public static implicit operator float(XLDate xDate)
        {
            return (float)xDate._xlDate;
        }

        public static implicit operator XLDate(double xlDate)
        {
            return new XLDate(xlDate);
        }

        public static implicit operator DateTime(XLDate xDate)
        {
            return DateTime.FromOADate(xDate);
        }

        public static implicit operator XLDate(DateTime dt)
        {
            return new XLDate(dt);
        }
        #endregion

        #region formatting

        public override string ToString()
        {
            return DateTime.FromOADate(_xlDate).ToString();
        }

        public string ToString(string format)
        {
            return DateTime.FromOADate(_xlDate).ToString(format);
        }

        public string ToString(string format,IFormatProvider formatprovider)
        {
            return DateTime.FromOADate(_xlDate).ToString(format,formatprovider);
        }

        #endregion

        #region System

        public override bool Equals(object obj)
        {
            if (obj is XLDate)
            {
                return ((XLDate)obj)._xlDate == _xlDate;
            }
            else if (obj is double)
            {
                return ((double)obj) == _xlDate;
            }
            else
                return false;
        }

        public override int GetHashCode()
        {
            return _xlDate.GetHashCode();
        }

        public int CompareTo(object target)
        {
            if (!(target is XLDate))
                throw new ArgumentException();

            return (this._xlDate).CompareTo(((XLDate)target)._xlDate);
        }

        #endregion


    }
}
