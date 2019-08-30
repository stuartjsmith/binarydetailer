using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace NetVersionChecker
{
    internal class Program
    {
        private static string currentBinary = string.Empty;

        private static void Main(string[] args)
         {
            string path = args[0];
            List<BinaryDetail> binaryDetails = new List<BinaryDetail>();
            Environment.CurrentDirectory = path;

            IEnumerable<string> allFiles = GetFiles(path, new[] { "*.dll", "*.exe" }, SearchOption.TopDirectoryOnly);
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += new ResolveEventHandler(CurrentDomain_ReflectionOnlyAssemblyResolve);

            foreach (string binary in allFiles)
            {
                currentBinary = binary;
                BinaryDetail bd = new BinaryDetail(binary);
                binaryDetails.Add(bd);

                try
                {
                    Assembly loadFrom = Assembly.LoadFrom(binary);

                    bd.ImageRuntimeVersion = loadFrom.ImageRuntimeVersion;
                    bd.AssemblyVersion = loadFrom.GetName().Version.ToString();
                    bd.FileVersion = FileVersionInfo.GetVersionInfo(binary).FileVersion;
                    bd.ProductVersion = FileVersionInfo.GetVersionInfo(binary).ProductVersion;

                    PortableExecutableKinds peKind;
                    ImageFileMachine machine;
                    loadFrom.ManifestModule.GetPEKind(out peKind, out machine);
                    bd.PEKind = peKind;
                    bd.ImageFileMachine = machine;

                    ProcessorArchitecture pe = loadFrom.GetName().ProcessorArchitecture;
                    bd.ProcessorArchitecture = pe;

                    Assembly loadFromReflectionOnly = Assembly.ReflectionOnlyLoadFrom(binary);
                    bool TargetFrameworkAttributeFound = false;
                    foreach (CustomAttributeData att in loadFromReflectionOnly.CustomAttributes)
                    {
                        if (att.AttributeType == typeof(System.Runtime.Versioning.TargetFrameworkAttribute))
                        {
                            bd.TargetFrameworkAttribute = att.ConstructorArguments[0].ToString();
                            TargetFrameworkAttributeFound = true;
                            break;
                        }
                    }
                    if(!TargetFrameworkAttributeFound)
                    {
                        string message = "No TargetFrameworkAttribute found on this assembly";
                        bd.Error = message;
                    }
                }
                catch (Exception e)
                {
                    bd.Error = e.Message;
                }
            }

            string fileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Guid.NewGuid() + ".csv");
            File.Create(fileName).Close();
            File.AppendAllLines(fileName, BinaryDetail.CSVHeader);
            foreach(var nfc in binaryDetails)
            {
                File.AppendAllLines(fileName, new[] { nfc.ToCsv() });
            }
            Console.WriteLine("File created at " + fileName);
            Console.In.ReadLine();
        }

        private static Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)
        {
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string assemblyPath = Path.GetDirectoryName(currentBinary) + "\\" + assemblyName.Name + ".dll";

            if (File.Exists(assemblyPath))
            {
                return Assembly.ReflectionOnlyLoadFrom(assemblyPath);
            }
            return Assembly.ReflectionOnlyLoad(args.Name);
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
