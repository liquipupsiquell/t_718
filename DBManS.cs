using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

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

    class CreateTable
    {
        public void StartProcessing(NonterminalExpression arg)//Arguments: 1)name database 2)name new table 3)-more is data for table
        {
            string[] args = arg.NoTInfo;
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
                return;
            }
            xlApp.Visible = true;
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;

            var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[xlSheets.Count], Type.Missing, Type.Missing, Type.Missing);//Создаем новый лист
           // xlSheets[xlSheets.Count - 1].Visible = false;                                                                                                   // string ValueForCells = "";

            for (int i = 0; i < args.Length-2; i++)
            {
                xlNewSheet.Cells[1, i + 1] = args[i+2];
                // xlNewSheet.Cells[1,i+1]=TableInfo[i];
            }
            xlNewSheet.Name = Convert.ToString(arg.NoTInfo[1]);//указываем имя книги
         //   xlNewSheet.Visible = true;
            //xlNewSheet.Range["A0"].TextToColumns("jhk");
            xlWorkBook.ReadOnlyRecommended = false;//выключаем защиту документа
            xlWorkBook.Save();
            xlWorkBook.Close();
            Console.WriteLine("Making table is end");

        }
        public void EndProcessing()
        {
            //действия по завершению создания таблицы
        }
    }

    class CreateTableCommand : ICommand
    {
        CreateTable tv;
        NonterminalExpression argm;
        public CreateTableCommand(CreateTable tvSet,NonterminalExpression argmore)
        {
            tv = tvSet;
            argm = argmore;
        }
        public void Execute()
        {
            tv.StartProcessing(argm);
        }
        public void Undo()
        {
            tv.EndProcessing();
        }
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
    class FuncSelect
    {
        public void StartProcessing()
        {
            //действия для функции вывода таблицы
        }

        public void StopProcessing()
        {
            //действия по завершению вывода таблицы
        }
    }
    class FuncSelectCommand : ICommand
    {
        TerminalExpression argOb;
        FuncSelect select;

        public FuncSelectCommand(FuncSelect m, int t, TerminalExpression arg)
        {
            argOb = arg;
            select = m;
        }
        public void Execute()
        {
            select.StartProcessing();
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
                Console.WriteLine("The name is occupied by another database!Try more... Or take select.function for change Table in DB");
                return;
            }
            else
            {
                xlWorkBook = xlApp.Workbooks.Add();
                Console.WriteLine("Database creation...");
            }
            xlWorkBook.SaveAs(@"D:\" + arg.NoTInfo[0] + ".xlsx");//сохраняем файл как...
            xlWorkBook.Close();

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
    ///</patternCommand>
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
            }//How terminal exp. have key for collection with objects commands
            //Nonterminal expression have arguments for processing data in this command
            

            FuncCreateDB createDB = new FuncCreateDB();//make object
            CreateTable createT = new CreateTable();
            Dictionary<string, ICommand> ListCommand = new Dictionary<string, ICommand>();//make collection
            ListCommand.Add("CreateDB", new FuncCreateDBCommand(createDB,ob2));//distribution him in collection and add key for him
            ListCommand.Add("CreateTable", new CreateTableCommand(createT, ob2));

            Set set = new Set();
            set.SetCommand(ListCommand[ob1.TerInfo]);
            set.Processing();

           // Console.ReadKey();
        }

    }
    ///</main>

}
