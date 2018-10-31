using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using PatternCommand;
namespace DBManS
{
    class Delete
    {
        public void StartProcessing(NonterminalExpression argm)
        {
            string[] arg = argm.NoTInfo;

            
            if (arg.Length == 2)
            {
                Console.WriteLine("Delete database...");

                try
                {
                    File.Delete(@"D:\" + argm.NoTInfo[0] + ".xlsx");
                }
                catch (IOException e)
                {
                    Console.WriteLine(e.Message);
                    return;
                }
            }
            else if (arg.Length == 3)
            {
                Console.WriteLine("Delete table...");
                

                try
                {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbook xlWorkBook;
                    xlApp.Visible = true;
                    xlWorkBook = xlApp.Workbooks.Open(@"D:\" + argm.NoTInfo[0] + ".xlsx");
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[arg[1]];
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int temp = lastCell.Row;
                    for (int j = 1,i=0; i!= temp;i++)//clearing table
                    {
                        Console.WriteLine(j+"||"+lastCell.Row);
                        
                            ObjWorkSheet.Rows[j].Delete();
                    }
                    xlWorkBook.Save();//Table is Clear

                    xlWorkBook = xlApp.Workbooks.Open(@"D:\" + argm.NoTInfo[0] + ".xlsx");
                    var sheet = (Excel.Worksheet)xlWorkBook.Sheets[arg[1]];//delete clear table
                    sheet.Delete();
                   xlWorkBook.Save();
                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message); Console.ReadKey();
                    return;
                }
            }
            else if (arg.Length == 4)
            {
                Console.WriteLine("Delete row...");
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbook xlWorkBook;
                try
                {
                    xlWorkBook = xlApp.Workbooks.Open(@"D:\" + argm.NoTInfo[0] + ".xlsx");
                   
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return;
                }
                Regex ob2 = new Regex(@"[id].[0-9]{0,255}");
                Regex ob1 = new Regex(@"[column].[A-Z||a-z||0-9]{0,12}.[A-Z||a-z||0-9]{0,12}");
                string[] args = argm.NoTInfo;
                for (int i = 1; i < args.Length - 1; i++)
                {
                    if (ob1.IsMatch(args[i])) ;
                    else if (ob2.IsMatch(args[2])) ;
                    else
                    {
                        Console.WriteLine("Incorrect syntax, try again or call the helpme command");
                        return;
                    }

                }
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[Convert.ToString(arg[1])];
                var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                if (ob1.IsMatch(arg[2]))
                {
                    string temp = arg[2], NameColumn = "" ;string type="";
                     type =temp.Remove(temp.IndexOf("."));
                    
                    Console.WriteLine(type);
                    if (type == "column")
                    {
                        Console.WriteLine("Processing..."+temp);
                        temp = temp.Remove(0, temp.IndexOf(".")+1);
                        NameColumn = temp.Remove(temp.IndexOf("."));
                        temp = temp.Remove(0, temp.IndexOf(".")+1);


                     //   string temp3 = temp.Remove(0, 1);
                        for (int i = 1; i < lastCell.Column; i++)
                        {
                         //   Console.WriteLine(ObjWorkSheet.Cells[1, i].Text.ToString() + "||" + temp + ".txt" + "||" + NameColumn + ".txt" + "||" + temp3 + ".txt");
                            if (ObjWorkSheet.Cells[1, i].Text.ToString() == NameColumn + ".txt")
                            {
                                Console.WriteLine("Найден столбец, ищем строку "+temp);
                                
                                for (int j = 2; j <= lastCell.Row; j++)
                                {
                                    Console.WriteLine(ObjWorkSheet.Cells[j, i].Text.ToString());
                                    if (ObjWorkSheet.Cells[j, i].Text.ToString() == temp)
                                    {
                                        Console.WriteLine("Найден строка");
                                        ObjWorkSheet.Rows[j].Delete();
                                        Console.WriteLine("Удаление строки произошло"); break;
                                    }
                                   
                                }
                                break;
                            }
                        }
                    }
                }
                    if (ob2.IsMatch(arg[2]))
                    {
                        
                        string RowId = arg[2];
                        RowId = RowId.Remove(0,RowId.IndexOf(".")+1);
                        ObjWorkSheet.Rows[RowId+1].Delete();
                    }

                xlWorkBook.Save();
                xlWorkBook.Close();
                ClearExcel.Clear_all();
                Console.WriteLine("End...");

            }
        }
    }
    class DeleteCommand : ICommand
    {
        Delete delete;
        NonterminalExpression argm;
        public DeleteCommand(Delete _delete, NonterminalExpression argmore)
        {
            delete = _delete;
            argm = argmore;
        }
        public void Execute()
        {
            delete.StartProcessing(argm);
        }
    }
}
