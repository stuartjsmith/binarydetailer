using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Versioning;
using System.Security.Policy;

namespace BinaryDetailer
{
    public class BinaryDetailFactory
    {
        public BinaryDetail CreateBinaryDetail(FileInfo binaryFileInfo)
        {
            if (string.IsNullOrEmpty(binaryFileInfo.Directory.FullName))
                throw new InvalidOperationException(
                    "Directory can't be null or empty.");

            if (!Directory.Exists(binaryFileInfo.Directory.FullName))
                throw new InvalidOperationException(
                    string.Format(CultureInfo.CurrentCulture,
                        "Directory not found {0}",
                        binaryFileInfo.Directory.FullName));

            var childDomain = BuildChildDomain(
                AppDomain.CurrentDomain);

            try
            {
                var bd = new BinaryDetail(binaryFileInfo);
                var loaderType = typeof(BinaryDetailPopulator);
                if (loaderType.Assembly != null)
                {
                    var loader =
                        (BinaryDetailPopulator) childDomain.CreateInstanceFrom(
                            loaderType.Assembly.Location,
                            loaderType.FullName).Unwrap();

                    try
                    {
                        loader.LoadAssembly(
                            binaryFileInfo.FullName);

                        bd = loader.PopulateBinaryDetail(bd);
                    }
                    catch (Exception e)
                    {
                        // this is the error thrown I think where we try to load a native dll using ReflectionOnlyLoadFrom, so ignore that specific message
                        if (!e.Message.Contains("The module was expected to contain an assembly manifest."))
                        {
                            bd.Error = e.Message;
                        }
                    }
                }

                return bd;
            }
            finally
            {
                AppDomain.Unload(childDomain);
            }
        }

        /// <summary>
        ///     Creates a new AppDomain based on the parent AppDomains
        ///     Evidence and AppDomainSetup
        /// </summary>
        /// <param name="parentDomain">The parent AppDomain</param>
        /// <returns>A newly created AppDomain</returns>
        private AppDomain BuildChildDomain(AppDomain parentDomain)
        {
            var evidence = new Evidence(parentDomain.Evidence);
            var setup = parentDomain.SetupInformation;
            return AppDomain.CreateDomain("DiscoveryRegion",
                evidence, setup);
        }

        private class BinaryDetailPopulator : MarshalByRefObject
        {
            [SuppressMessage("Microsoft.Performance",
                "CA1822:MarkMembersAsStatic")]
            internal BinaryDetail PopulateBinaryDetail(BinaryDetail binaryDetail)
            {
                var directory = new DirectoryInfo(binaryDetail.FileInfo.DirectoryName);

                ResolveEventHandler resolveEventHandler =
                    (s, e) =>
                    {
                        return OnReflectionOnlyResolve(
                            e, directory);
                    };

                AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve
                    += resolveEventHandler;

                var reflectionOnlyAssembly =
                    AppDomain.CurrentDomain.ReflectionOnlyGetAssemblies().First();

                binaryDetail.ImageRuntimeVersion = reflectionOnlyAssembly.ImageRuntimeVersion;
                binaryDetail.AssemblyVersion = reflectionOnlyAssembly.GetName().Version.ToString();

                PortableExecutableKinds peKind;
                ImageFileMachine machine;
                reflectionOnlyAssembly.ManifestModule.GetPEKind(out peKind, out machine);
                binaryDetail.PEKind = peKind;
                binaryDetail.ImageFileMachine = machine;

                var pe = reflectionOnlyAssembly.GetName().ProcessorArchitecture;
                binaryDetail.ProcessorArchitecture = pe;

                binaryDetail.TargetFrameworkAttribute = GetCustomAttributeValue(reflectionOnlyAssembly, typeof(TargetFrameworkAttribute));

                AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve
                    -= resolveEventHandler;

                return binaryDetail;
            }

            private string GetCustomAttributeValue(Assembly assembly, Type type)
            {
                var attributesList = assembly.CustomAttributes
                    .Where(a => a.AttributeType == type).ToList();
                if (attributesList.Count == 1)
                    return attributesList[0].ConstructorArguments[0].ToString().Replace("\"", "");
                return string.Empty;
            }

            private Assembly OnReflectionOnlyResolve(
                ResolveEventArgs args, DirectoryInfo directory)
            {
                var loadedAssembly =
                    AppDomain.CurrentDomain.ReflectionOnlyGetAssemblies()
                        .FirstOrDefault(
                            asm => string.Equals(asm.FullName, args.Name,
                                StringComparison.OrdinalIgnoreCase));

                if (loadedAssembly != null) return loadedAssembly;

                var assemblyName =
                    new AssemblyName(args.Name);
                var dependentAssemblyFilename =
                    Path.Combine(directory.FullName,
                        assemblyName.Name + ".dll");

                if (File.Exists(dependentAssemblyFilename))
                    return Assembly.ReflectionOnlyLoadFrom(
                        dependentAssemblyFilename);
                return Assembly.ReflectionOnlyLoad(args.Name);
            }

            [SuppressMessage("Microsoft.Performance",
                "CA1822:MarkMembersAsStatic")]
            internal void LoadAssembly(string assemblyPath)
            {
                try
                {
                    Assembly.ReflectionOnlyLoadFrom(assemblyPath);
                }
                catch (FileNotFoundException)
                {
                    /* Continue loading assemblies even if an assembly
                     * can not be loaded in the new AppDomain. */
                }
            }
        }
    }
}