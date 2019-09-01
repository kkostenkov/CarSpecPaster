using System;
using System.Collections.Generic;
using System.IO;

namespace CarSpecPaster
{
    class Program
    {
        public class Options
        {
            private const string TextFilesDirectoryName = "TextFiles";
            private const string RulesFileName = "ReplacementRules.txt";
            private const string TemplateFileName = "Template.doc";
            private const string OutputFileName = "Result.doc";
            

            public string RulesFilePath = "D:\\Repos\\zTemp\\ReplacementRules.txt";
            public string XlsFilePath = "D:\\Repos\\zTemp\\List_Microsoft_Excel.xlsx";
            public string OriginFilePath = "D:\\Repos\\zTemp\\zakl.doc";
            public string ResultFilePath = "D:\\Repos\\zTemp\\zaklResult.doc";

            public Options()
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                Console.WriteLine(baseDir);
                var textFilesDir = Path.Combine(baseDir, TextFilesDirectoryName);
                RulesFilePath = Path.Combine(textFilesDir, RulesFileName);
                OriginFilePath = Path.Combine(textFilesDir, TemplateFileName);
                ResultFilePath = Path.Combine(Environment.CurrentDirectory, OutputFileName);
                if (!FindXls())
                {
                    Environment.Exit(0);
                }
            }

            private bool FindXls()
            {
                var localFiles = Directory.EnumerateFiles(Environment.CurrentDirectory);
                foreach (var fileName in localFiles)
                {
                    if (fileName.EndsWith(".xls") || fileName.EndsWith(".xlsx"))
                    {
                        XlsFilePath = Path.Combine(Environment.CurrentDirectory, fileName);
                        return true;
                    }
                }
                Console.WriteLine(string.Format("Excel file not found in {0}", Environment.CurrentDirectory));
                return false;
            }
        }

        
        static void Main(string[] args)
        {
            var options = new Options();
            var rulesReader = new RulesReader();
            var rules = rulesReader.Read(options.RulesFilePath);

            var xlsReader = new XlsReader();

            var excelData = new Dictionary<string, string>();
            foreach (var request in rules.Requests)
            {
                var data = xlsReader.Read(options.XlsFilePath, request);
                foreach (var kvp in data)
                {
                    excelData[kvp.Key] = kvp.Value;
                }
            }

            var substitutor = new StringSubstitutor();
            var substitutions = substitutor.ComposeSubstitutionPairs(rules, excelData);
            substitutor.Work(options.OriginFilePath, options.ResultFilePath, substitutions);
        }
    }
}
