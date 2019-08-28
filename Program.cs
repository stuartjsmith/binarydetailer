using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NetVersionChecker
{
    class Program
    {
         static void Main(string[] args)
        {

            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += new ResolveEventHandler(CurrentDomain_ReflectionOnlyAssemblyResolve);
            string path = args[0];
            Environment.CurrentDirectory = path;
            foreach(string binary in MyDirectory.GetFiles(path, "\\.dll|\\.exe", SearchOption.AllDirectories))
            {
                Console.WriteLine(binary);
                try
                {
                    Assembly loadFrom = Assembly.LoadFrom(binary);
                    Console.WriteLine("ImageRuntimeVersion: " + loadFrom.ImageRuntimeVersion);

                    Assembly loadFromReflectionOnly = Assembly.ReflectionOnlyLoadFrom(binary);
                    bool TargetFrameworkAttributeFound = false;
                    foreach (CustomAttributeData att in loadFromReflectionOnly.CustomAttributes)
                    {
                        if (att.AttributeType == typeof(System.Runtime.Versioning.TargetFrameworkAttribute))
                        {
                            Console.WriteLine("TargetFrameworkAttribute:" + att.ConstructorArguments[0].ToString());
                            TargetFrameworkAttributeFound = true;
                            break;
                        }
                    }
                    if(!TargetFrameworkAttributeFound)
                    {
                        Console.WriteLine("No TargetFrameworkAttribute found on this assembly");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                Console.WriteLine();
            }
        }

        static Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)

        {

            return System.Reflection.Assembly.ReflectionOnlyLoad(args.Name);

        }
    }

    public static class MyDirectory
    {   // Regex version
        public static IEnumerable<string> GetFiles(string path,
                            string searchPatternExpression = "",
                            System.IO.SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            Regex reSearchPattern = new Regex(searchPatternExpression, RegexOptions.IgnoreCase);
            return Directory.EnumerateFiles(path, "*", searchOption)
                            .Where(file =>
                                     reSearchPattern.IsMatch(Path.GetExtension(file)));
        }

        // Takes same patterns, and executes in parallel
        public static IEnumerable<string> GetFiles(string path,
                            string[] searchPatterns,
                            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            return searchPatterns.AsParallel()
                   .SelectMany(searchPattern =>
                          Directory.EnumerateFiles(path, searchPattern, searchOption));
        }
    }
}
