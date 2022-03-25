﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

                foreach (var row in sheetData.Elements<Row>())
                {
                    foreach (var cell in row.Elements<Cell>())
                    {
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;

                            Console.WriteLine("Shared string {0}: {1}", ssid, str);
                        }
                        else
                        {
                            Console.WriteLine("Shared string {0}: {1}", );
                        }
                    }
                }

                //foreach (Cell cell in cells)
                //{
                //    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                //    {
                //        int ssid = int.Parse(cell.CellValue.Text);
                //        string str = sst.ChildElements[ssid].InnerText;
                //        Console.WriteLine("Shared string {0}: {1}", ssid, str);
                //    }
                //    else if (cell.CellValue != null)
                //    {
                //        Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                //    }
                //}

                //foreach (Row row in rows)
                //{
                //    foreach (Cell c in row.Elements<Cell>())
                //    {
                //        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                //        {
                //            int ssid = int.Parse(c.CellValue.Text);
                //            string str = sst.ChildElements[ssid].InnerText;
                //            Console.WriteLine("Shared string {0}: {1}", ssid, str);
                //        }
                //        else if (c.CellValue != null)
                //        {
                //            Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                //        }
                //    }
                //}
            }
        }
    }
}