using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace helpme
{
    class OutInfo
    {
        public void OutListCommands()
        {
            //creating database
            Console.WriteLine("CreateDB-this command have functions for make DataBase\n" +
                "Input agruments:\nCreateDB <name_for_new_Database> - make new database with the specified name\n");
            //creating new table in DB
            Console.WriteLine("CreateTable-this command make new table\nInput arguments:\n" +
                "CreateTable <name_database> <name_for_new_table> agrument1 argument2 argument3 arument4 - make table in specified database\n");
            //Select (switch) 
            Console.WriteLine("Select-this command used for out information about DB or out table on board\n" +
                "Input arguments:\nSelect <name_database> - out list of all tables in this database\n" +
                "Select <name_database> <name_table> - out this table in console\n");
            //Update/change information in table
            Console.WriteLine("Update-this command help you add new info in table\n" +
                "Update <name_database> <name_table> (1,2)=text1 (4,2)=text2 (1,4)=text3\n");
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
    }
    class Program
    {
        static void Main(string[] args)
        {
            OutInfo inf = new OutInfo();
            inf.OutListCommands();
        }
    }
}
