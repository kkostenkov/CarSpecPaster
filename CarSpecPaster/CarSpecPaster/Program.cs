using System;
using System.Collections.Generic;

namespace CarSpecPaster
{
    class Program
    {

        private static string rulesFilePath = "D:\\Repos\\zTemp\\ReplacementRules.txt";
        private static string xlsFilePath = "D:\\Repos\\zTemp\\List_Microsoft_Excel.xlsx";

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var rulesReader = new RulesReader();
            var rules = rulesReader.Read(rulesFilePath);

            var xlsReader = new XlsReader();

            var dict = new Dictionary<string, string>();
            foreach (var request in rules.Requests)
            {
                var data = xlsReader.Read(xlsFilePath, request);
                foreach (var kvp in data)
                {
                    dict[kvp.Key] = kvp.Value;
                }
            }
        }
    }
}
