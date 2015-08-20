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
        private Lazy<string[]> _colnames;
        private Lazy<string[]> _propnames;
        private int _columns;

        private Action<object, int, object> _setters;
        private Func<object, int, object> _getters;
        
        #endregion

        #region constructors

        public XLObjectMapping(Type t, Func<IXLObjectMapping> factory = null)
        {
            _t = t;

            if (factory == null && !typeof(IXLObjectMapping).IsAssignableFrom(t))
            {
                var fieldinfos = t.GetProperties();

                var propnames = fieldinfos.Select(f => f.Name).ToArray();

                SetColnames(t, propnames, propnames);
            }
            else
            {
                IXLObjectMapping instance = (factory != null) ? factory() : (t.GetConstructor(Type.EmptyTypes) != null) ? (IXLObjectMapping)Activator.CreateInstance(t, new object[0])
                            : (IXLObjectMapping)Activator.CreateInstance(t);
                _columns = instance.ColumnCount();

                _setters = (o, i, v) => ((IXLObjectMapping)o).SetColumn(i, v);
                _getters = (o, i) => ((IXLObjectMapping)o).GetColumn(i);

                _colnames = new Lazy<string[]>(() =>
                {
                    var retval = new string[_columns];
                    for (int i = 0; i < _columns; i++)
                        retval[i] = instance.ColumnName(i);
                    return retval;
                });
                _propnames = _colnames;

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

            _columns = colnames.Length;

            var setters = propnames.Select(p => new Action<object, object>((o, v) =>
            {
                var f = t.GetProperty(p);
                f.SetValue(o, v.ConvertTo(f.PropertyType), null);
            })).ToArray();
            _setters = (o, i, rhs) =>
            {
                if (i >= 0 && i < _columns)
                    setters[i](o, rhs);
            };
            var getters = propnames.Select(p => new Func<object, object>(o => t.GetProperty(p).GetValue(o, null))).ToArray();
            _getters = (o, i) =>
            {
                if (i >= 0 && i < _columns)
                    return getters[i](o);
                else
                    return null;
            };
            _colnames = new Lazy<string[]>(() => colnames);
            _propnames = new Lazy<string[]>(() => propnames);

        }

        #endregion

        #region properties

        public Type MappedType
        {
            get { return _t; }
        }

        public string[] Colnames
        {
            get { return _colnames.Value; }
        }

        public string[] Propnames
        {
            get { return _propnames.Value; }
        }

        public int Columns
        {
            get { return _columns; }
        }


        #endregion

        #region access to getters and setters

        public object GetColumn(object instance, int index)
        {
            if (index >= 0 && index < _columns)
                return _getters(instance, index);
            else
                return null;
        }

        public void SetColumn(object instance, int index, object RHS)
        {
            if (index >= 0 && index < _columns)
            _setters(instance, index, RHS);
        }


        
        #endregion
    }
}
