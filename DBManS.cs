using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
namespace PatternComannd
{
    /// <interpret>
    class Context
    {
        string[] strComand;
        public Context(string[] value)
        {
            strComand = value;
        }
        public string[] info
        {
            get { return strComand; }
        }

    }

    abstract class AbstractExpression
    {
        public abstract void Interpret(Context context);
    }
    class TerminalExpression : AbstractExpression
    {
        private string str;
        public string TerInfo
        {
            get
            {
                return str;
            }
        }
        public override void Interpret(Context context)
        {
            string[] a = context.info;
            str = a[0];
        }
    }
    class NonterminalExpression : AbstractExpression
    {
        string[] str;
        public string[] NoTInfo
        {
            get
            {
                return str;
            }
        }
        public override void Interpret(Context context)
        {
            str = new string[context.info.Length];
            for (int i = 1; i < context.info.Length; i++)
            {
                str[i - 1] = context.info[i];
            }
        }
    }
    /// </interpret>

    ///<patternCommand>
    interface ICommand
    {
        void Execute();
        void Undo();
    }

    // Invoker
    class Set
    {
        ICommand command;

        public Set() { }

        public void SetCommand(ICommand com)
        {
            command = com;
        }

        public void Processing()
        {
            command.Execute();
        }
        public void EndProcessing()
        {
            command.Undo();
        }
    }
    class CreateTable
    {
        public void StartProcessing(NonterminalExpression arg)//Arguments: 1)name database 2)name new table 3)-more is data for table
        {
            Regex ob1 = new Regex(@"[A-Z||a-z]{1}[a-z||0-9]{0,12}.[int||txt||date]");
            string[] args = arg.NoTInfo;
            for(int i = 2; i < args.Length-1; i++)
            {

                if (ob1.IsMatch(args[i]));
                else
                {
                    Console.WriteLine("Incorrect syntax, try again or call the helpme command");
                    return;
                }

            }


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;
            FileInfo fi = new FileInfo("D:\\" + arg.NoTInfo[0] + ".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            if (fi.Exists)
            {
                Console.WriteLine("File is open, processing...");
                xlWorkBook = xlApp.Workbooks.Open(@"D:\" + arg.NoTInfo[0] + ".xlsx");

            }
            else
            {
                Console.WriteLine("DataBase {0} is not, try more...", arg.NoTInfo[0]);
                ClearExcel.Clear_all();
                return;
            }
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;

            var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[xlSheets.Count], Type.Missing, Type.Missing, Type.Missing);//Создаем новый лист
           // xlSheets[xlSheets.Count - 1].Visible = false;                                                                                                   // string ValueForCells = "";

            for (int i = 0; i < args.Length-2; i++)
            {
                xlNewSheet.Cells[1, i + 1] = args[i+2];
                // xlNewSheet.Cells[1,i+1]=TableInfo[i];
            }
            xlNewSheet.Name = Convert.ToString(arg.NoTInfo[1]);//указываем имя книги
            //xlNewSheet.Visible = true;
            //xlNewSheet.Range["A0"].TextToColumns("jhk");
            xlWorkBook.ReadOnlyRecommended = false;//выключаем защиту документа
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();
            ClearExcel.Clear_all();
            Console.WriteLine("Making table is end");
            Console.ReadKey();
        }
        public void EndProcessing()
        {
            //действия по завершению создания таблицы
        }
    }

    class CreateTableCommand : ICommand
    {
        CreateTable create;
        NonterminalExpression argm;
        public CreateTableCommand(CreateTable m,NonterminalExpression argmore)
        {
            create = m;
            argm = argmore;
        }
        public void Execute()
        {
            create.StartProcessing(argm);
        }
        public void Undo()
        {
            create.EndProcessing();
        }
    }
    //Select - command for out table at console
    class FuncSelect//last change
    {
        public void StartProcessing(NonterminalExpression arg)
        {
            string[] arguments = arg.NoTInfo;

            Regex ob1 = new Regex(@"[A-Z||a-z]{1}[a-z||0-9]{0,12}");
            string[] args = arg.NoTInfo;
            for (int i = 1; i < args.Length - 1; i++)
            {
                if (ob1.IsMatch(args[i]));
                else
                {
                    Console.WriteLine("Incorrect syntax, try again or call the helpme command");
                    return;
                }

            }
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo("D:\\" + arg.NoTInfo[0] + ".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            if (fi.Exists)
            {
                xlWorkBook = xlApp.Workbooks.Open(@"D:\" + arg.NoTInfo[0] + ".xlsx");
            }
            else
            {
                Console.WriteLine("File is not!");
                return;
            }
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            if (arguments.Length == 2)
            {
                string[] result;

                result = new string[xlWorkBook.Sheets.Count];
                try
                {
                    var sheet = (Excel.Worksheet)xlWorkBook.Sheets["Лист1"];//delete default sheet
                    sheet.Delete();
                }catch(Exception e) { }
                try
                {
                    for (int i = 0; i < xlWorkBook.Sheets.Count; i++)
                    {

                        Console.WriteLine((i + 1) + ". " + ((Excel.Worksheet)xlWorkBook.Sheets[i + 1]).Name);//out list of names tables
                        xlWorkBook.Save();
                       // ClearExcel.Clear_all();
                    }
                }
                catch (Exception e) { }

            }
            else
            {
                if (arguments.Length == 3)
                {
                    Console.WriteLine(Convert.ToString(arguments[1]));
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[Convert.ToString(arguments[1])]; //получить 1 лист
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                    string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                    for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                        for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                            list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
                    for (int i = 0; i < lastCell.Row; i++) {
                        for (int j = 0; j < lastCell.Column; j++) {
                            if (list[j, i].IndexOf(".") !=-1) list[j, i] = list[j, i].Remove(list[j, i].IndexOf("."));
                            Console.Write("{0,-10}|",list[j, i]);
                        }Console.WriteLine();
                    }
                }
            }
            xlWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ClearExcel.Clear_all();
        }

        public void StopProcessing()
        {
            //действия по завершению вывода таблицы
        }
    }
    class FuncSelectCommand : ICommand
    {
        NonterminalExpression argOb;
        FuncSelect select;

        public FuncSelectCommand(FuncSelect m, NonterminalExpression arg)
        {
            argOb = arg;
            select = m;
        }
        public void Execute()
        {
            select.StartProcessing(argOb);
            select.StopProcessing();
        }

        public void Undo()
        {
            select.StopProcessing();
        }
    }
    /////CreateDB
    class FuncCreateDB
    {
        public void StartProcessing(NonterminalExpression arg)
        {
            Console.WriteLine("Start creating file for DB");

            
            //действия для функции оздания базы данных
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo("D:\\" + arg.NoTInfo[0] + ".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            if (fi.Exists)
            {
                Console.WriteLine("The name is occupied by another database!Try more... Or take select function for change Table in DB");
                return;
            }
            else
            {
                xlWorkBook = xlApp.Workbooks.Add();
                Console.WriteLine("Database creation...");
            }
            xlWorkBook.SaveAs(@"D:\" + arg.NoTInfo[0] + ".xlsx");//сохраняем файл как...
            ClearExcel.Clear_all();
        }
        public void StopProcessing()
        {
            Console.WriteLine("Ready, you can continue self job (for list of command enter \"helpme\")");
        }
    }
    class FuncCreateDBCommand : ICommand
    {
        NonterminalExpression argOb;
        FuncCreateDB create;
        public FuncCreateDBCommand(FuncCreateDB m, NonterminalExpression arg)
        {
            create = m;
            argOb = arg;
        }
        public void Execute()
        {
            create.StartProcessing(argOb);
            create.StopProcessing();
        }

        public void Undo()
        {
            create.StopProcessing();
        }
    }
    ///</AddInfo>
class Update
{
      public void StartProcessing(NonterminalExpression arg)
      {
            string[] arguments = arg.NoTInfo;
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();//Подклюаем Excel
            Excel.Workbook xlWorkBook;
            //FileInfo fi = new FileInfo("D:\\" + arg.NoTInfo[0] + ".xlsx");//Проверяем есть ли файл с таким же названием, если да то выведит ошибку, т.к. БД уже сществует
            //if (fi.Exists)
            //{
            //    xlWorkBook = xlApp.Workbooks.Open(@"D:\" + arg.NoTInfo[0] + ".xlsx");
            //}
            //else
            //{
            //    Console.WriteLine("File is not!");
            //    return;
            //}
            
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(@"D:\" + arg.NoTInfo[0] + ".xlsx");
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);
                return;
            }
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[Convert.ToString(arguments[1])]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();//считываем текст в строку

            string tempstr = "";string TypeText = "";
            int ii = lastCell.Row+1, jj = 1;
            int check = 0;
                foreach (string k in arguments)
                {
                if (k == arguments[0] || k == arguments[1]) continue;
                try
                {
                    
                    if (jj > lastCell.Column) { Console.WriteLine("Error syntax"); return; }
                    if (k.IndexOf("[") != -1)
                    {
                      TypeText= taketype(ObjWorkSheet.Cells[1, jj].Text.ToString());
                        Regex ob1 = new Regex(TypeText);
                        tempstr = k.Remove(k.IndexOf("["), 1);
                        if (ob1.IsMatch(tempstr))
                        ObjWorkSheet.Cells[ii, jj] = tempstr;
                        else
                        {
                            Console.WriteLine("according't types...");return;
                        }
                        jj++;
                    }
                    else if (k.IndexOf("]") != -1)
                    {
                        TypeText = taketype(ObjWorkSheet.Cells[1, jj].Text.ToString());

                        tempstr = k.Remove(k.IndexOf("]"), 1);
                        Regex ob2 = new Regex(TypeText);
                        if(ob2.IsMatch(tempstr))
                        ObjWorkSheet.Cells[ii , jj ] = tempstr;
                        else
                        {
                            Console.WriteLine("according't types..."); return;
                        }
                        ii++; jj = 1; 
                    }
                    else
                    {
                        TypeText = taketype(ObjWorkSheet.Cells[1, jj].Text.ToString());

                        Regex ob3 = new Regex(TypeText);
                        if (ob3.IsMatch(k))
                            ObjWorkSheet.Cells[ii, jj] = k;
                        else
                        {
                            Console.WriteLine("according't types..."); return;
                        }

                        jj++;
                        
                    }
                }
                catch (Exception e) { Console.WriteLine(e.Message); }
                }
            xlWorkBook.Save();
        }
        public string taketype(string a)
        {
            a = a.Remove(0, a.IndexOf('.'));
            if (a == ".int") a = @"[0-9]";
            if (a == ".txt") a = @"[A-z]";
            if (a == ".date") a = @"[0-9]{1,2}.[0-9]{1,2}.[0-9]{4}";
            return a;
        }
    public void EndProcessing()
    {
    }
        public string INT(string inputData)
        {
            Regex ob1 = new Regex(@"[0-9]{0,10}");
                if (ob1.IsMatch(inputData));
                else
                {
                    Console.WriteLine("Incorrect type data in {0}...",inputData);
                    return "-666";
                }



            return inputData;
        }
        public void TXT()
        {

        }
}

class UpdateCommand : ICommand
{
    Update Upd;
    NonterminalExpression argm;
    public UpdateCommand(Update m, NonterminalExpression argmore)
    {
        Upd = m;
        argm = argmore;
    }
    public void Execute()
    {
        Upd.StartProcessing(argm);
    }
    public void Undo()
    {
        Upd.EndProcessing();
    }
}
/////


///</patternCommand>
class OutInfo:ICommand
    {
        public void Execute()
        {
            //creating database
            Console.WriteLine("CreateDB-this command have functions for make DataBase\n" +
                "Input agruments:\nCreateDB <name_for_new_Database> - make new database with the specified name\n");
            //creating new table in DB
            Console.WriteLine("CreateTable-this command make new table\nInput arguments:\n" +
                "CreateTable <name_database> <name_for_new_table> agrument1.typesdata argument2.typesdata argument3.typesdata arument4.typesdata - make table in specified database\n");
            //Select (switch) 
            Console.WriteLine("Select-this command used for out information about DB or out table on board\n" +
                "Input arguments:\nSelect <name_database> - out list of all tables in this database\n" +
                "Select <name_database> <name_table> - out this table in console\n");
            //Update/change information in table
            Console.WriteLine("Update-this command help you add new info in table\n" +
                "Update <name_database> <name_table> [text1 text2 text3] [text4 text5 text6] etc.\n");
            //Delete info of id and etc.
            Console.WriteLine("Delete-this command can to delete column or row by id_numbers or title column\nInput arguments:\n" +
                "Delete <name_database> - delete database\n" +
                "Delete <name_database> <name_table> - delete table from this database\n" +
                "Delete <name_database> <name_table> <name_column/number_row> - delete full row or column by input name-argument\n");
            //далее будут выводиться возможные запросы, пока что только команды для работы с БД
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
            ////
            //Console.WriteLine("");
        }
        public void Undo()
        {

        }
    }
    class ClearExcel
    {
        public static void Clear_all()
        {
            string nameproc = "Excel";
            System.Diagnostics.Process[] etc = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process anti in etc)
            {
                try
                {
                    if (anti.ProcessName.ToLower().Contains(nameproc.ToLower())) anti.Kill();
                }
                catch(Exception e)
                {

                }
            }
        }
    }
    ///<main>
    class Program
    {
        static void Main(string[] args)
        {
            //работаем с аргументами в интерпретаторе
            var context = new Context(args);
            var list = new List<AbstractExpression>();
            TerminalExpression ob1 = new TerminalExpression();
            NonterminalExpression ob2 = new NonterminalExpression();
            list.Add(ob1);
            list.Add(ob2);
            foreach (AbstractExpression exp in list)
            {
                exp.Interpret(context);
            }//terminal exp. have key for collection with objects commands
            //Nonterminal expression have arguments for processing data in this command
            

            FuncCreateDB createDB = new FuncCreateDB();//make object
            CreateTable createT = new CreateTable();
            FuncSelect select = new FuncSelect();
            Update Upd = new Update();
            Dictionary<string, ICommand> ListCommand = new Dictionary<string, ICommand>();//make collection
            ListCommand.Add("helpme", new OutInfo());
            ListCommand.Add("CreateDB", new FuncCreateDBCommand(createDB,ob2));//distribution him in collection and add key for him
            ListCommand.Add("CreateTable", new CreateTableCommand(createT, ob2));
            ListCommand.Add("Select", new FuncSelectCommand(select, ob2));
            ListCommand.Add("Update", new UpdateCommand(Upd, ob2));

            Set set = new Set();
            set.SetCommand(ListCommand[ob1.TerInfo]);
            set.Processing();
            ClearExcel.Clear_all();
            Console.ReadKey();

        }

    }
    ///</main>

}
