using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// CMD command: BinaryDetailer.exe "C:\Program Files\dotnet\sdk\6.0.400\cs" doc

namespace BinaryDetailer
{
    internal class Program
    {
        private static List<string> ExcludeNames = new List<string>();
        
        private static void Main(string[] args)
        {
            List<BinaryDetail> groupedStrings = new List<BinaryDetail>();

            string path = args[0];
            if (args.Length > 0)
            {
                for (int idx = 1; idx < args.Length; idx++)
                {
                    ExcludeNames.Add(args[idx].ToLower());
                }
            }

            Export ex = new Export(ExcludeNames);

            List<BinaryDetail> binaryDetails = new List<BinaryDetail>();
            IEnumerable<string> allFiles = GetFiles(path, new[] { "*.dll", "*.exe" }, SearchOption.AllDirectories);

            foreach (string binary in allFiles)
            {
                BinaryDetail bd = new BinaryDetailFactory().CreateBinaryDetail(new FileInfo(binary));
                binaryDetails.Add(bd);
            }

            ex.CreateReport(binaryDetails);

            if (args.Length > 1 && args[1].ToLower().Equals("doc"))
            {
                ex.CreateWordDoc(binaryDetails);
            }

            ex.GroupBinary(binaryDetails, @"C:\AVEVA\GIT\binarydetailer\GroupingConfigExample.xml");

            ex.CreateReport(binaryDetails);

            Console.WriteLine("Export complete");
            Console.In.ReadLine();
        }

        private static IEnumerable<string> GetFiles(string path,
            string[] searchPatterns,
            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            return searchPatterns.AsParallel()
                .SelectMany(searchPattern =>
                    Directory.EnumerateFiles(path, searchPattern, searchOption));
        }
    }
}