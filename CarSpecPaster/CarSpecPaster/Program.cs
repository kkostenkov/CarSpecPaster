using System;
using System.Collections.Generic;

namespace CarSpecPaster
{
    class Program
    {

        private static string rulesFilePath = "D:\\Repos\\zTemp\\ReplacementRules.txt";
        private static string xlsFilePath = "D:\\Repos\\zTemp\\List_Microsoft_Excel.xlsx";
        private static string originFilePath = "D:\\Repos\\zTemp\\zakl.doc";
        private static string resultFilePath = "D:\\Repos\\zTemp\\zaklResult.doc";

        static void Main(string[] args)
        {
            var rulesReader = new RulesReader();
            var rules = rulesReader.Read(rulesFilePath);

            var xlsReader = new XlsReader();

            var excelData = new Dictionary<string, string>();
            foreach (var request in rules.Requests)
            {
                var data = xlsReader.Read(xlsFilePath, request);
                foreach (var kvp in data)
                {
                    excelData[kvp.Key] = kvp.Value;
                }
            }

            var substitutor = new StringSubstitutor();
            var substitutions = substitutor.ComposeSubstitutionPairs(rules, excelData);
            substitutor.Work(originFilePath, resultFilePath, substitutions);
        }
    }
}
