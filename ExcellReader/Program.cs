// See https://aka.ms/new-console-template for more information
using ExcelDataReader;

using System.Collections.Generic;
using Spire.Xls;
using Spire.Xls.Core;


Console.WriteLine("Hello, World!");
string filePath = @"C:\Users\huseyin.ergun\Desktop\newSEO.xlsx";

////Dosyayı okuyacağımı ve gerekli izinlerin ayarlanması.
//FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
////Encoding 1252 hatasını engellemek için;

//System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

//IExcelDataReader excelReader;
//List<string> newliste = new List<string>();
//List<string> newliste2 = new List<string>();
//int counter = 0;

////Gönderdiğim dosya xls'mi xlsx formatında mı kontrol ediliyor.
//if (Path.GetExtension(filePath).ToUpper() == ".XLS")
//{
//    //Reading from a binary Excel file ('97-2003 format; *.xls)
//    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
//}
//else
//{
//    //Reading from a OpenXml Excel file (2007 format; *.xlsx)
//    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
//}
//string alpha = "zxcvbnmasdfghjkliqwertyuop-1234567890";
//string list;
//string list2;
////Veriler okunmaya başlıyor.
//while (excelReader.Read())
//{
//    list = excelReader.GetString(0).Trim();
//    //list2 = excelReader.GetString(1).Trim();
//    foreach(var l in list)
//    {
//        if (!alpha.Contains(l))
//        {
//            list=list.Replace(l,'-').Replace("--","-").Replace("---","-").Replace("----", "-").Replace("-----", "-").Trim('-');
//        }
//        newliste.Add(list);
//    }
    
//}

//Console.WriteLine("İlk liste veri sayısı " + newliste.Count + " İkinci liste veri sayısı " + newliste2.Count);
//excelReader.Close();

string filePath2 = "C:\\Users\\huseyin.ergun\\Desktop\\newSEO.xlsx";

// Create a Workbook object
Workbook workbook = new Workbook();
//Remove default worksheets
//workbook.Worksheets.Clear();
//Add a worksheet and name it
//Worksheet worksheet = workbook.Worksheets.Add("WriteToCell");
workbook.LoadFromFile(filePath2);

Worksheet worksheet = workbook.Worksheets["Sheet2"];
//Write data to specific cells
worksheet.Range[1, 1].Value = "Student Name";
worksheet.Range[1, 2].Value = "Math";
worksheet.Range[1, 3].Value = "English";
worksheet.Range[1, 4].Value = "Total Marks";
worksheet.Range[2, 1].Value = "Hazel";
worksheet.Range[2, 2].NumberValue = 80;
worksheet.Range[2, 3].NumberValue = 78;
worksheet.Range[2, 4].NumberValue = 158;
worksheet.Range[3, 1].Value = "Tina";
worksheet.Range[3, 2].NumberValue = 98;
worksheet.Range[3, 3].NumberValue = 72;
worksheet.Range[3, 4].NumberValue = 170;
//Auto fit column width
worksheet.AllocatedRange.AutoFitColumns();
//Save to an Excel file
workbook.SaveToFile("C:\\Users\\huseyin.ergun\\Desktop\\newSEO.xlsx");


