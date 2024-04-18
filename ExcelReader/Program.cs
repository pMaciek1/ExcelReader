using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int highScore = 0;
            bool init = true;
            int score;
            string highScoreName="";
            FileInfo fi = new FileInfo("database.db");
            try
            {
                if (fi.Exists)
                {
                    using (var connection = new SQLiteConnection("Data Source=database.db"))
                    {
                        connection.Open();
                        fi.Delete();
                        Console.WriteLine("Deleting existing database file...");
                    }
                }
            }
            catch
            {
                fi.Delete();
                Console.WriteLine("Deleting existing database file...");
            }
            Console.WriteLine("Creating new database file...");
            SQLiteConnection.CreateFile("database.db");
            using(var connection = new SQLiteConnection("Data Source=database.db"))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = "Create Table scoreTable (Name TEXT, Score INT)";
                command.ExecuteNonQuery();
                connection.Close();
            }
            Console.WriteLine("Created new database file.");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage("C:\\Users\\Maciek\\source\\repos\\ExcelReader\\ExcelReader\\Scores.xlsx"))
            {
                var sheet = package.Workbook.Worksheets["Scores"];
                DataTable dataTable = sheet.Cells["A2:B16"].ToDataTable();
                foreach (DataRow row in dataTable.Rows)
                {
                    Console.WriteLine("---------------------------------");
                    Console.Write(row[0].ToString() + ": ");
                    Console.Write(row[1].ToString());
                    Console.Write('\n');
                    Console.WriteLine("---------------------------------");
                    score = int.Parse(row[1].ToString());
                    if (init==true || score > highScore)
                    {
                        highScoreName = row[0].ToString();
                        highScore = score;
                        init = false;
                    }
                }
            }
            Console.WriteLine("Highest Score: " + highScore + " - " + highScoreName);
            Console.ReadKey();
        }
    }
}
