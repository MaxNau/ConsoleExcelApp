using System;

namespace ConcoleExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Pleas enter table size:\n ");
            Console.Write("Pleas enter number of rows: ");
            var rows = int.Parse(Console.ReadLine());
            Console.Write("Pleas enter number of columns: ");
            var columns = int.Parse(Console.ReadLine());

            var table = new Table(rows, columns);
            table.PopulateTable();
            table.CalculateResult();
            table.DisplayTable(true);
            Console.ReadKey();
        }
    }
}
