ExcelDna.Utilities Ver. 0.1
==========================

Utilities that add functionality to ExcelDna such as creating a COM like interface to the C API

This work builds on the excellent ExcelDna library by Govert van Drimmelen (http://excel-dna.net/).

You can get the latest build from nuget. While the library is still in development, great care is taken to not break backward compatibility.


- Light weight and intuitive (at least for myself) to use, trying to stay close to COM interface but adding C# language specifics such as generics, lambdas, collections


- Have Workbook, Worksheet and other object types that nicely wrap the functionality (eg. Workbook.SaveAs())

- Range utilities like .ToVector<T> or .ToMatrix<T> that automatically do the correct type conversions (also for enums, Dates)

- A double based date type XLDate equivalent to DateTime accelerating interaction with math libraries

- Triggering / or preventing Recalculation by passing Action<> (or a lambda), the typical Screenupdating = false pattern

- Interacting with sheets / range contents as List<T> where T is a row object similar to an ORM

- DataTable extensions for interacting with excel ranges

Excel is a trademark of Microsoft



## General application functions


```csharp

using ExcelDna.Integration;
using ExcelDna.Utilities;

public static class TestMacros
{

	[ExcelCommand]
	public static void Test_click()
	{
	
		// message bar
		XLApp.MessageBar("This is number {0}", 5);

		// set calculation
		XLApp.Calcuation = xlCalculation.Manual;
		XLApp.CalculateNow();
		XLApp.CalculateDocument();

		// execute something on a range 
		// can be on a different sheet; will return
		XLApp.ActionOnSelectedRange(r, () => { });
		
		XLApp.ReturnToSelection(() => { });

		// suspend screen updating and calculation
		Action a = ()=>{ /* lots of cell updating etc. */};
		a.NoCalcAndUpdating();  //extension method


		//All open workbooks
		Workbook[] workbooks = XLApp.Workbooks;

		// Copy & paste
		XLApp.SelectRange("rangeRef");
		XLApp.Copy();
		XLApp.PasteSpecial(xlPasteType.PasteAll, 
			xlPasteAction.None, skip_blanks: false, transpose: false);

	}
}
```
		
## Workbooks and Worksheets

```csharp


	[ExcelCommand]
	public static void Test1_click()
	{
		// get active work sheet; usually the easiest way to get reference
		Worksheet ws = Worksheet.ActiveSheet();
		
		// get the workbook
		Workbook wb = ws.Workbook;

		// [Workbook]SheetName
		string sheetref = ws.SheetRef;

		// constructors
		// does not create sheet in workbook
		var ws1 = new Worksheet("[workbook]Sheetname"); 
		var ws2 = new Worksheet(wb1, "sheet1");
		
		var wb1 = new Workbook(@"filepath");

		// creates sheet in workbook
		Worksheet ws3 = wb1.AddWorksheet();
		ws3.Name = "Sheet3";

		//Workbook functions
		wb1.Activate();
		wb1.Save();
		wb1.SaveAs("filepath", password: "", write_password: "", read_only: false);
		wb1.Close(saveChanges: true, routeFile: false);
		string wbpath = wb1.GetPath();

		// {"sheet1","sheet2"}
		string[] sheetnames = wb1.SheetNames;
		// {"[Workbook1]sheet1","[Workbook1]sheet2"}
		string[] sheetrefs = wb1.SheetRefs;

		Worksheet[] sheets = wb1.Worksheets;
		// "Workbook1.xlsx"
		string wbname = wb1.Name;
		// "c:\temp\Workbook1.xlsx"
		string path = wb1.Path;
	}

```

## Ranges

```csharp

	[ExcelCommand]
	public static void Test2_click()
	{

		// Get a range object
		ExcelReference range = Name.GetRange("[Workbook]Sheet!Name");

		// copy & delete
		range.ClearContents(); //delete values
		range.Copy(toRange: null);
		range.DeleteEntireRows(); //shift cells up
		
		// offsetting & resizing
		var r = range.Offset(1, 1);
		r = r.Resize(10, 10);
		string referesto = r.RefersTo(); //"=R1C2:R2C5"

		// formatting
		r.FormatBorder(/* lots of options*/);
		r.FormatColor(backcolor: 0, forecolor: 0, pattern: 0);
		r.FormatNumber("YYYY.MM.DD");
		// use this to figure out default date format for your version of excel
		string defaultDateFormat = XLApp.DefaultDateFormat;
		r.FormatNumber(defaultDateFormat);

		// get values
		string val = range.GetValue<string>();
		double[] vec = range.GetValue().ToVector<double>();
		double[,] mat = range.GetValue().ToMatrix<double>();
	}

```


## Names

```csharp

	[ExcelCommand]
	public static void Test3_click()
	{
		// get active work sheet; usually the easiest way to get reference
		Worksheet ws = Worksheet.ActiveSheet();
		var wb = ws.Workbook;
		
		// create a local name
		var name1 = ws.AddName("name1", "=R1C1:R5C5", hidden: false, local: true);

		bool isglobal = name1.IsGlobalScope;
		bool isLocal = name1.IsLocalScope;
		string nameLocal = name1.NameLocal;
		// "[Workbook]Sheet!Name" - or - "Workbook!Name"
		string nameRef = name1.NameRef;
		//"=R1C1:R5C5"
		string x = name1.RefersTo;

		Name[] names_all = ws.Names; //local and workbook
		Name[] names_local = ws.NamesLocal;

		Name.GetValue<string>("[Workbook]Sheet!Name");
		Name.GetValue<string>(ws, "localname");
		Name.GetValue<string>(wb, "globalname");

		// Get a range object
		ExcelReference range = Name.GetRange("[Workbook]Sheet!Name");
	}

```

## Interaction with List<T> and DataTable

There exist a simple object mapper that allows to interact with ranges and strongly typed lists.

```csharp

    class Person
    {
	
		/*
		Mark properties that should be ignored by the mapper with XLIgnorePropertyAttribute
		
		It is also possible to use your own IgnorePropertyAttribute
		
		Then set XLObjectMapping.IgnorePropertyAttribute = typeof(MyIgnorePropertyAttribute); at startup
		This has the obvious advantage that domain model objects don't need a reference to this assembly
		*/
	
		[XLIgnoreProperty]
		public long DBID { get; set; }
		
        public string Name { get; set; }
        public string Address { get; set; }
        public DateTime BirthDay { get; set; }
        public int Age { get; set; }

    }

	[ExcelCommand]
	public static void Test4_click()
	{
		// get active work sheet; usually the easiest way to get reference
		Worksheet ws = Worksheet.ActiveSheet();
		var wb = ws.Workbook;
		
		// if the range has the same order of fields than the object we
		// can do the following
		List<Person> persons = ws.Range("persons").ToList<Person>();

		// and dump back
		// this will automatically adjust the size of the named output range
		ws.Range("persons").Fill(persons, "persons", header: false);
	}

```

This simple way to interact with POCOs will only work for simple field types (string, DateTime, double, ...) and if the order of the properties in the class is the same is in the columns. However, we can influence how this mapping can be done.

Another way to map the properties is to define the mapping directly in the object mapper. 

```csharp

	//Get only Name and Adress fields and use MyName and MyAdress for the header instead
	XLObjectMapper.SetObjectMapping(new XLObjectMapping(typeof(Person),
		new string[] { "MyName", "MyAddress" }, new string[] {"Name","Adress" }));

```

If we want to control more granularly how to map the fields we can implement the following interface. This is particularly useful if certain fields are arrays or class types. The object mapper will automatically pick this interface up.

```csharp
    /// <summary>
    /// Interface that allows to enumerate and name the fields of an object
    /// typically used for interacting with excel ranges and Lists of strongly typed objects
    /// that cannot be enumerated easily or automatically through reflection
    /// XLObjectMapping will use these values instead of reflection if the type implements this interface
    /// </summary>
    public interface IXLObjectMapping
    {
        /// <summary>
        /// Number of columns
        /// </summary>
        /// <returns></returns>
        int ColumnCount();
        /// <summary>
        /// Column name for index 
        /// </summary>
        /// <param name="index">zero based</param>
        /// <returns></returns>
        string ColumnName(int index);
        /// <summary>
        /// Indexed getter
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        object GetColumn(int index);
        /// <summary>
        /// Indexed setter
        /// </summary>
        /// <param name="index"></param>
        /// <param name="RHS"></param>
        void SetColumn(int index, object RHS);
    }
```
