using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace OpenXMLExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\Izagakhmaevra\Desktop\Excel\TestExel.xlsx";

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    static Category GetCategory(SheetData sheetData, SharedStringTable sst)
                    {
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            foreach (var cell in row.Elements<Cell>())
                            {
                                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                {
                                    int ssid = int.Parse(cell.CellValue.Text);
                                    string str = sst.ChildElements[ssid].InnerText;

                                    //Console.WriteLine("Shared string {0}: {1}", ssid, str);

                                    Category category = new Category();

                                    category.name = str;
                                    category.type = str;
                                    category.format = str;
                                    category.laung = str;
                                     
                                    return category;
                                }
                                //else
                                //{
                                //    //Console.WriteLine("Shared string {0}: {1}", );
                                //}

                                static void Print(Category category)
                                {
                                    Console.WriteLine($"ID: {category.id}");
                                    Console.WriteLine($"Имя: {category.name}");
                                    Console.WriteLine($"Тип файла: {category.type}");
                                    Console.WriteLine($"Формат файла: {category.format}");
                                    Console.WriteLine($"Язык: {category.laung}");
                                }

                                static void Additional(string[] args, Category firstCategory)
                                {
                                    Print(firstCategory);
                                }
                            }
                        }
                    }                   
                }
            }
        }
    }
}
