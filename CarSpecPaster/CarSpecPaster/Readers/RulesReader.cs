using System;
using System.Collections.Generic;
using System.IO;

namespace CarSpecPaster
{
    public class ReadRules
    {
        internal List<XlsReadRequest> Requests;
        internal Dictionary<string, string> Substitutions;
    }

    class RulesReader
    {
        public ReadRules Read(string filePath)
        {
            var fileExists = File.Exists(filePath);
            if (!fileExists)
            {
                Console.WriteLine($"No file found at {0}", filePath);
                return null;
            }
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var sr = new StreamReader(stream);
                var readRequest = ReadRules(sr);
                return readRequest;
            }
        }

        private ReadRules ReadRules(StreamReader sr)
        {
            var rules = new ReadRules();
            Console.WriteLine("Reading rules...");
            rules.Requests = ReadRequests(sr);
            rules.Substitutions = ReadSubstitutions(sr);
            return rules;
        }

        private const string COMMENT = @"//";
        private const string SHEET_NAME = @"Лист=";
        private const string RULES_SECTION_NAME = @"#Правила";
        private const string KV_SEPARATOR = @"=";

        private List<XlsReadRequest> ReadRequests(StreamReader sr)
        {
            var requests = new List<XlsReadRequest>();
            XlsReadRequest request;
            do
            {
                request = ReadRequestBlock(sr);
                if (request != null)
                {
                    requests.Add(request);
                }
            }
            while (request != null);

            return requests;
        }

        private XlsReadRequest ReadRequestBlock(StreamReader sr)
        {
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line.StartsWith(COMMENT))
                {
                    continue;
                }

                if (line.StartsWith(RULES_SECTION_NAME))
                {
                    return null;
                }

                if (line.StartsWith(SHEET_NAME))
                {
                    var settingLine = line.Substring(SHEET_NAME.Length);
                    var splitSettings = settingLine.Split(",");
                    if (splitSettings.Length != 3)
                    {
                        Console.WriteLine($"I don't know how to parse setting line {0}", settingLine);
                        return null;
                    }
                    var request = new XlsReadRequest();
                    request.SheetName = splitSettings[0];
                    request.KeysColumn = Int32.Parse(splitSettings[1]);
                    request.ValuesColumn = Int32.Parse(splitSettings[2]);
                    Console.WriteLine(string.Format("From Excel sheet {0} read columns {1} as keys and {2} as values:", 
                        request.SheetName, request.KeysColumn, request.ValuesColumn));
                    return request;
                }

            }
            return null;
        }

        private Dictionary<string, string> ReadSubstitutions(StreamReader sr)
        {
            var substitutions = new Dictionary<string, string>();
            Console.WriteLine("Replacement rules:" );
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line.StartsWith(COMMENT))
                {
                    continue;
                }

                if (!line.Contains(KV_SEPARATOR))
                {
                    continue;
                }
                var kv = line.Split(KV_SEPARATOR);
                substitutions[kv[0]] = kv[1];
                Console.WriteLine(string.Format("{0}:{1}", kv[0], kv[1]));

            }
            return substitutions;
        }
    }
}
