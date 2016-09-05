using System;

namespace ConcoleExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Pleas, enter table size: ");
            //Console.Write("Enter number of rows: ");

            var rowsColumns = Console.ReadLine();

            // Split user input string into seprate strings that correspond to row cells
            while (string.IsNullOrWhiteSpace(rowsColumns))
            {
                rowsColumns = Console.ReadLine();
            }

            int rows;
            int columns;

            var rowColNum = rowsColumns.Split(' ');
            while (rowColNum.Length < 2 || rowColNum.Length > 2 | !int.TryParse(rowColNum[0], out rows) | !int.TryParse(rowColNum[1], out columns))
            {
                rowsColumns = Console.ReadLine();
                if (rowsColumns != null) rowColNum = rowsColumns.Split(' ');
            }

            /*while (!int.TryParse(Console.ReadLine(), out rows))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Must be an integer");
                Console.ResetColor();
                Console.Write("Enter number of rows: ");
            }

            Console.Write("Enter number of columms: ");
            int columns;

            while (!int.TryParse(Console.ReadLine(), out columns))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Must be an integer");
                Console.ResetColor();
                Console.Write("Enter number of columns: ");
            }*/

            var table = new Table(rows, columns);

            table.PopulateTable();
            table.CalculateResult();
            table.DisplayTable(true);
            Console.ReadKey();
        }
    }
}
