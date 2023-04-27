using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

            if (args.Length > 1)
            {
                foreach (var arg in args.Select((value, i) => new { i, value }))
                {
                    var index = arg.i;

                    // Skip the first arg as this is the path to report on
                    if (index.Equals(0)) continue;

                    if (arg.value.ToLower().Equals("doc"))
                    {
                        // Create word document
                        ex.CreateWordDoc(binaryDetails);
                    }
                    else if (arg.value.ToLower().Equals("config"))
                    {
                        // Create a binary grouping based on a config.
                        ex.GroupBinary(binaryDetails, args[index+1]);
                        ex.CreateReport(binaryDetails);
                    }
                }
            }
            else
            {
                // Raw report
                ex.CreateReport(binaryDetails);
            }

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