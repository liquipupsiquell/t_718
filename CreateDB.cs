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
            //args = new string[0];
            //args[0] = "nameDB [id nameclient activity]";
            string name = " ";string str= "NameDateBase [id NameClient Activity]";//Для теста, в последующем будет из args
            int j = 0;
           
           // str = args[0];
            foreach(char k in str)//узнаем имя базы данных
            {
                if (k == ' ') { str = str.Remove(0, str.IndexOf(' ')+1); break; }
                name += k;

            }

            int amount=1;//узнаем количество столбцов
            foreach (char k in str)
            {
                if (k == ' ') { amount++; }
            }

            string[] TableInfo=new string[amount];
            for (int i = 0; i < amount; i++)//Обрабатывает args 
            {
                if (str.IndexOf(' ')+1 > 0)
                {
                    TableInfo[i] = str.Substring(0, str.IndexOf(' ')+1);
                    if (TableInfo[i].First() == '[') TableInfo[i] = TableInfo[i].Remove(0, 1);
                    str = str.Remove(0, str.IndexOf(' ') + 1);
                }
                else
                {
                    TableInfo[i] = str.Substring(0, str.IndexOf(']'));
                    str = str.Remove(0, str.IndexOf(']') + 1);
                }
            }
            


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo("D:\\"+name+".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            if (fi.Exists)
            {
                xlWorkBook = xlApp.Workbooks.Open(@"D:\"+name+".xlsx");
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
            for(int i = 0;i<amount; i++)
            {   
                    xlNewSheet.Cells[1,i+1]=TableInfo[i];
            }

            xlNewSheet.Name = Convert.ToString(name);//указываем имя книги
            //xlNewSheet.Range["A0"].TextToColumns("jhk");
            xlWorkBook.ReadOnlyRecommended = false;//выключаем защиту документа
            if (fi.Exists) xlWorkBook.Save();
            else xlWorkBook.SaveAs(@"D:\" + name + ".xlsx");//сохраняем файл как...
            xlWorkBook.Close();
        }
    }
}
