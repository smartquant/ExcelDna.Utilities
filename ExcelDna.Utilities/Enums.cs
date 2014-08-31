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
    public enum xlCalculation
    {
        Automatic = 1,
        SemiAutomatic = 2,
        Manual = 3,
    }

    public enum xlPasteType
    {
        PasteAll = 1,
        PasteFormulas = 2,
        PasteValues = 3,
        PasteFormats = 4,
        PasteNotes = 5
    }

    public enum xlPasteAction
    {
        None = 1,
        Add = 2,
        Substract = 3,
        Multiply = 4,
        Divide = 5,
    }

    public enum xlUpdateLinks
    {
        Never = 0,
        ExternalOnly = 1,
        RemoteOnly = 2,
        ExternalAndRemote = 3,
    }

    public enum xlBorderStyle
    {
        NoBorder = 0,
        ThinLine = 1,
        MediumLine = 2,
        DashedLine = 3,
        DottedLine = 4,
        ThickLine = 5,
        Doubleline = 6,
        HairLine = 7,
    }
}
