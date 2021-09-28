using RootLogic;
using System;
using System.IO;

namespace ExcelDiffs
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("First file path: ");
            var firstFilePath = Console.ReadLine().Trim();
            Console.Write("Column number: ");
            var firstFileColumnNumber = int.Parse(Console.ReadLine().Trim());

            Console.Write("First file path: ");
            var secondFilePath = Console.ReadLine().Trim();
            Console.Write("Column number: ");
            var secondFileColumnNumber = int.Parse(Console.ReadLine().Trim());

            var service = new LogicService(firstFilePath, firstFileColumnNumber, secondFilePath, secondFileColumnNumber);
            var resultsPath = Path.Combine(Directory.GetCurrentDirectory(), "result.txt");
            service.DiffAndSave(resultsPath);

            Console.WriteLine("Done, homie!");
        }
    }
}
