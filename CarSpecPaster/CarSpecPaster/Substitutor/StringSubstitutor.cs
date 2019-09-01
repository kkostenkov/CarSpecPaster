using Spire.Doc;
using System;
using System.Collections.Generic;

namespace CarSpecPaster
{
    class StringSubstitutor
    {
        public void Work(string originFilePath, string resultFilePath, 
            Dictionary<string, string> substitutions)
        {
            Document doc = new Document();
            doc.LoadFromFileInReadMode(originFilePath, FileFormat.Auto);
            //doc.LoadFromFile(originFilePath);
            Replace(doc, substitutions);
            doc.SaveToFile(resultFilePath, FileFormat.Doc);
        }

        public Dictionary<string, string> ComposeSubstitutionPairs(ReadRules rules, Dictionary<string, string> excelData)
        {
            var result = new Dictionary<string, string>();
            Console.WriteLine("Composed pairs:");
            foreach (var item in rules.Substitutions)
            {
                var ruleKey = item.Key;
                var excelName = item.Value;
                if (excelData.ContainsKey(excelName))
                {
                    result[ruleKey] = excelData[excelName];
                    Console.WriteLine(string.Format("{0} : {1}", ruleKey, result[ruleKey]));
                }
            }
            return result;
        }

        private void Replace(Document doc, Dictionary<string, string> substitees)
        {
            foreach (var kvp in substitees)
            {
                var replacesCount = doc.Replace(kvp.Key, kvp.Value, true, true);
                if (replacesCount < 1)
                {
                    Console.WriteLine($"Failed to find key {0} to replace it with {1}", kvp.Key, kvp.Value);
                }
            }
        }
    }
}
