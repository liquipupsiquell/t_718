using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel =Microsoft.Office.Interop.Excel;
using System.IO;

namespace CreateNewDB
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo("D:\\"+args[0]+".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            if (fi.Exists)
            {
                xlWorkBook = xlApp.Workbooks.Open(@"D:\"+ args[0] + ".xlsx");
                Console.WriteLine("The name is occupied by another database!Try more... Or take select.function for change Table in DB");
                xlWorkBook.ReadOnlyRecommended = false;
                xlWorkBook.Close();
                return;
            }
            else
            {
                xlWorkBook = xlApp.Workbooks.Add();
                Console.WriteLine("Database creation...");
            }
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;// 
            
           var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[xlSheets.Count], Type.Missing, Type.Missing, Type.Missing);//Создаем новый лист
           // string ValueForCells = "";

            for(int i = 0;i<args.Length-2; i++)
            {
                xlNewSheet.Cells[1, i + 1] = args[i + 2];
                   // xlNewSheet.Cells[1,i+1]=TableInfo[i];
            }
            xlNewSheet.Name = Convert.ToString(args[1]);//указываем имя книги
            //xlNewSheet.Range["A0"].TextToColumns("jhk");
            xlWorkBook.ReadOnlyRecommended = false;//выключаем защиту документа
            if (fi.Exists) xlWorkBook.Save();
            else xlWorkBook.SaveAs(@"D:\" + args[0] + ".xlsx");//сохраняем файл как...
            xlWorkBook.Close();
            Console.WriteLine("Ready");
        }
    }
}
