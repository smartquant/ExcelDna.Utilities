ExcelDna.Utilities Ver. 0.1
==========================

Utilities that add functionality to ExcelDna such as creating a COM like interface to the C API

This work builds on the excellent ExcelDna library by Govert van Drimmelen (http://excel-dna.net/).

This is a early development build and still in development.


- Light weight and intuitive (at least for myself) to use, trying to stay close to COM interface but adding C# language specifics such as generics, lambdas, collections


- Have Workbook, Worksheet and other object types that nicely wrap the functionality (eg. Workbook.SaveAs())

- Range utilities like .ToVector<T> or .ToMatrix<T> that automatically do the correct type conversions (also for enums, Dates)

- A double based date type XLDate equivalent to DateTime accelerating interaction with math libraries

- Triggering / or preventing Recalculation by passing Action<> (or a lambda), the typical Screenupdating = false pattern

- Interacting with tables as List<T> where T is a row object similar to an ORM

- DataTable extensions for interacting with excel ranges


Todo's:

- Helpful examples



Excel is a trademark of Microsoft
