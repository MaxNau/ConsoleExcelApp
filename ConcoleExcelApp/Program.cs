using System;
using System.Diagnostics.CodeAnalysis;

namespace ConcoleExcelApp
{
    class Program
    {
        [SuppressMessage("ReSharper", "UnusedParameter.Local")]
        static void Main(string[] args)
        {
            Console.WriteLine("Pleas, enter table size: ");
            Console.Write("Enter number of rows: ");

            int rows;

            while (!int.TryParse(Console.ReadLine(), out rows))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Must be an integer");
                Console.ResetColor();
                Console.Write("Enter number of rows: ");
            }

            int columns;

            while (!int.TryParse(Console.ReadLine(), out columns))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Must be an integer");
                Console.ResetColor();
                Console.Write("Enter number of columns: ");
            }

            var table = new Table(rows, columns);

            table.PopulateTable();
            table.CalculateResult();
            table.DisplayTable(true);
            Console.ReadKey();
        }
    }
}
