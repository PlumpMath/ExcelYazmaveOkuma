using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelYazmaveOkuma
{
    class Program
    {
        static void Main(string[] args)
        {
            //Interop kullanımı ile projenin çalışması için bilgisayarda office olmalı.
            var excelApp = new Excel.Application();
            var dosyaYolu = @"C:\DB\test.xlsx";
            Excel.Workbook xlWorkbook = excelApp.Workbooks.Open(dosyaYolu, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    string temp = (xlRange.Cells[i, j] as Excel.Range).Value2 ?? "Boş";
                    Console.Write(temp+" ");
                    //Bunları ben ekranda gösterdim hücre hücre okuyor ve ekranda gösteriyor. Siz bir listeye atıp oradan istediğiniz şekilde oynama yaparak tekrar yazdırabilirsiniz. Tekrar yazdırmak için aşağıda yorum satırları içerisine aldığım yöntemi kullanabilirsiniz. 
                }
                Console.WriteLine();
            }

            Console.ReadLine();

            //Kayıt için oluşturacağımız excel dosyasının arka planda çalışmasını sağlar. true olursa görünür yapar. 
            //excelApp.Visible = false;

            //excelApp.Workbooks.Add();
            //Excel._Worksheet workSheet = (Excel.Worksheet) excelApp.ActiveSheet;
            //Burada koordinatlar veriliyor 1. satır A isimli sütun. 
            //workSheet.Cells[1, "A"] = "ID Number";
            //1. satır B isimli sütun
            //workSheet.Cells[1, "B"] = "Current Balance";
            //var bankAccounts=new List<Account>
            //{
            //    new Account
            //    {
            //        ID = 35234,
            //        Balance = 23424
            //    },
            //    new Account
            //    {
            //        ID = 234234234,
            //        Balance = -127.44
            //    }
            //};

            //var row = 1;
            //foreach (var acc in bankAccounts)
            //{
            //    row++;
            //    workSheet.Cells[row, "A"]=acc.Id;
            //    workSheet.Cells[row, "B"]=acc.Balance;
            //}
            //Sütunların genişlikliklerini ayarlıyor
            //workSheet.Columns[1].AutoFit();
            //workSheet.Columns[2].AutoFit();
            //test.xlsx olarak kaydediyor. Varsayılan olarak Belgelerim içine kaydeder. 
            //workSheet._SaveAs("test.xlsx");
        }
    }
    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }
}
