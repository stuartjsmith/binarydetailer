using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace BinaryDetailer
{
    internal class Program
    {
        private static string currentBinary = string.Empty;

        private static void Main(string[] args)
         {
            string path = args[0];
            List<BinaryDetail> binaryDetails = new List<BinaryDetail>();
            IEnumerable<string> allFiles = GetFiles(path, new[] { "*.dll", "*.exe" }, SearchOption.AllDirectories);

            foreach (string binary in allFiles)
            {
                BinaryDetail bd = new BinaryDetailFactory().CreateBinaryDetail(new FileInfo(binary));
                binaryDetails.Add(bd);
            }

            CreateReport(binaryDetails);
            Console.In.ReadLine();
         }

        private static void CreateReport(List<BinaryDetail> binaryDetails)
        {
            string fileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                Guid.NewGuid() + ".csv");
            File.Create(fileName).Close();
            File.AppendAllLines(fileName, BinaryDetail.CSVHeader);
            foreach (var nfc in binaryDetails)
            {
                File.AppendAllLines(fileName, new[] {nfc.ToCsv()});
            }

            Console.WriteLine("File created at " + fileName);
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
