using System.IO;
using NPOI.XSSF.UserModel;
using static System.Console;

var path = "/Users/gpdoud/Repos/TestNPOI/ExcelDoc.xlsx";

//XSSFWorkbook wb1 = null;
var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite);
var wb1 = new XSSFWorkbook(file);
file.Dispose();

wb1.GetSheetAt(0).GetRow(0).GetCell(0).SetCellValue("Sample");

var file2 = new FileStream(path, FileMode.Create, FileAccess.ReadWrite);
wb1.Write(file2);
file2.Close();
file2.Dispose();


WriteLine("Done ...");