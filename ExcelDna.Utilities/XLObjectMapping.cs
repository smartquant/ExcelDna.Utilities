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
using System.Linq;
using System.Text;


namespace ExcelDna.Utilities
{

    public class XLObjectMapping
    {
        #region fields

        private Type _t;
        private string[] _colnames;
        private string[] _propnames;

        //private string _separatorArrays = "|";

        private Action<object, object>[] _setters;
        private Func<object, object>[] _getters;
        
        #endregion

        #region constructors

        public XLObjectMapping(Type t)
        {
            _t = t;
            var mapinterface = t.GetInterfaces().FirstOrDefault(i => i == typeof(IXLObjectMapping));

            if (mapinterface == null)
            {
                var fieldinfos = t.GetProperties();

                var propnames = fieldinfos.Select(f => f.Name).ToArray();

                SetColnames(t, propnames, propnames);
            }
            else
            {
                IXLObjectMapping instance = (t.GetConstructor(Type.EmptyTypes) != null) ? (IXLObjectMapping)Activator.CreateInstance(t,new object[0]) 
                            : (IXLObjectMapping)Activator.CreateInstance(t);
                int cols = instance.ColumnCount();

                _colnames = new string[cols];
                _propnames = new string[cols];
                _setters = new Action<object, object>[cols];
                _getters = new Func<object, object>[cols];

                for (int i = 0; i < cols; i++)
                {
                    _colnames[i] = instance.ColumnName(i);
                    _propnames[i] = _colnames[i];
                    int j = i; //need to capture the variable
                    _setters[i] = new Action<object, object>((o, v) => ((IXLObjectMapping)o).SetColumn(j,v));
                    _getters[i] = new Func<object, object>(o => ((IXLObjectMapping)o).GetColumn(j));
                }
            }
        }

        public XLObjectMapping(Type t, string[] colnames, string[] propnames)
        {
            SetColnames(t, colnames, propnames);
        }

        private void SetColnames(Type t, string[] colnames, string[] propnames)
        {
            if (colnames == null || propnames == null) throw new ArgumentException("colnames == null || propnames == null!");
            if (colnames.Length != propnames.Length) throw new ArgumentException("colnames and propnames must have same length!");

            _setters = propnames.Select(p => new Action<object, object>((o, v) =>
            {
                var f = t.GetProperty(p);
                f.SetValue(o, v.ConvertTo(f.PropertyType), null);
            })).ToArray();
            _getters = propnames.Select(p => new Func<object, object>(o => t.GetProperty(p).GetValue(o, null))).ToArray();

            _colnames = colnames;
            _propnames = propnames;

        }

        #endregion

        #region properties

        public Type MappedType
        {
            get { return _t; }
        }

        public string[] Colnames
        {
            get { return _colnames; }
        }

        public string[] Propnames
        {
            get { return _propnames; }
        }

        public int Columns
        {
            get { return _colnames.Length; }
        }

        //public string SeparatorArrays
        //{
        //    get { return _separatorArrays; }
        //    set { _separatorArrays = value; }
        //}

        #endregion

        #region access to getters and setters

        public object GetColumn(object instance, int index)
        {
            if (index >= 0 && index < _colnames.Length)
                return _getters[index](instance);
            else
                return null;
        }

        public void SetColumn(object instance, int index, object RHS)
        {
            if (index >= 0 && index < _colnames.Length)
            _setters[index](instance, RHS);
        }


        
        #endregion
    }
}
