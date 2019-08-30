using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace NetVersionChecker
{
    class Program
    {
         static void Main(string[] args)
         {
            string path = args[0];
            List<NetFileCheck> netFileChecks = new List<NetFileCheck>();
            Environment.CurrentDirectory = path;

            IEnumerable<string> allFiles = GetFiles(path, new[] { "*.dll", "*.exe" }, SearchOption.TopDirectoryOnly);
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += new ResolveEventHandler(CurrentDomain_ReflectionOnlyAssemblyResolve);
            foreach(string binary in allFiles)
            {
                NetFileCheck nfc = new NetFileCheck(binary);
                netFileChecks.Add(nfc);
                Console.WriteLine(binary);
                try
                {
                    Assembly loadFrom = Assembly.LoadFrom(binary);
                    Console.WriteLine("ImageRuntimeVersion: " + loadFrom.ImageRuntimeVersion);
                    nfc.ImageRuntimeVersion = loadFrom.ImageRuntimeVersion;

                    Assembly loadFromReflectionOnly = Assembly.ReflectionOnlyLoadFrom(binary);
                    bool TargetFrameworkAttributeFound = false;
                    foreach (CustomAttributeData att in loadFromReflectionOnly.CustomAttributes)
                    {
                        if (att.AttributeType == typeof(System.Runtime.Versioning.TargetFrameworkAttribute))
                        {
                            Console.WriteLine("TargetFrameworkAttribute:" + att.ConstructorArguments[0].ToString());
                            nfc.TargetFrameworkAttribute = att.ConstructorArguments[0].ToString();
                            TargetFrameworkAttributeFound = true;
                            break;
                        }
                    }
                    if(!TargetFrameworkAttributeFound)
                    {
                        string message = "No TargetFrameworkAttribute found on this assembly";
                        Console.WriteLine(message);
                        nfc.Error = message;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    nfc.Error = e.Message;
                }
                Console.WriteLine();
            }
            string fileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Guid.NewGuid() + ".csv");
            File.Create(fileName).Close();
            File.AppendAllLines(fileName, new[] { "Full Path,FileName,ImageRuntimeVersion,TargetFrameworkAttribute,Error" });
            foreach(var nfc in netFileChecks)
            {
                File.AppendAllLines(fileName, new[] { nfc.ToCsv() });
            }
            Console.WriteLine("File created at " + fileName);
            Console.In.ReadLine();
        }

        static Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)
        {
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string assemblyPath = Environment.CurrentDirectory + "\\" + assemblyName.Name + ".dll";

            if (File.Exists(assemblyPath))
            {
                return Assembly.ReflectionOnlyLoadFrom(assemblyPath);
            }
            return Assembly.ReflectionOnlyLoad(args.Name);
        }

        static IEnumerable<string> GetFiles(string path,
                            string[] searchPatterns,
                            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            return searchPatterns.AsParallel()
                   .SelectMany(searchPattern =>
                          Directory.EnumerateFiles(path, searchPattern, searchOption));
        }
    }

    class NetFileCheck
    {
        public NetFileCheck(string fullPath)
        {
            FullPath = fullPath;
            FileName = Path.GetFileName(fullPath);
        }
        public string FullPath;
        public string FileName;
        public string ImageRuntimeVersion = String.Empty;
        public string TargetFrameworkAttribute = String.Empty;
        public string Error = String.Empty;

        public string ToCsv()
        {
            return FullPath + "," + FileName + "," + ImageRuntimeVersion + "," + TargetFrameworkAttribute + "," + Error;
        }
    }
}
