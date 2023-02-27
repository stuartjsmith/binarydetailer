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
        private static List<string> ExcludeNames = new List<string>();

        private static void Main(string[] args)
        {
            string path = args[0];
            if (args.Length > 0)
            {
                for (int idx = 1; idx < args.Length; idx++)
                {
                    ExcludeNames.Add(args[idx].ToLower());
                }
            }

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
            foreach (var binaryDetail in binaryDetails)
            {
                bool include = true;
                foreach (string excludeName in ExcludeNames)
                {
                    if ((binaryDetail.AssemblyCompanyAttribute != null && binaryDetail.AssemblyCompanyAttribute.ToLower().Contains(excludeName)) ||
                        (binaryDetail.AssemblyCopyrightAttribute != null && binaryDetail.AssemblyCopyrightAttribute.ToLower().Contains(excludeName)))
                    {
                        include = false;
                    }
                }

                if (include == false) continue;

                File.AppendAllLines(fileName, new[] { binaryDetail.ToCsv() });
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
