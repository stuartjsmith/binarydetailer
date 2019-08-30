using System;
using System.IO;
using System.Reflection;

namespace BinaryDetailer
{
    internal class BinaryDetail
    {
        public BinaryDetail(string fullPath)
        {
            FullPath = fullPath;
            FileName = Path.GetFileName(fullPath);
        }

        public string FullPath;
        public string FileName;
        public string ImageRuntimeVersion = String.Empty;
        public string TargetFrameworkAttribute = String.Empty;
        public string Error = String.Empty;
        public PortableExecutableKinds PEKind;
        public ImageFileMachine ImageFileMachine;
        public ProcessorArchitecture ProcessorArchitecture;

        public string ToCsv()
        {
            return 
                "\"" + FullPath + "\"," +
                "\"" + FileName + "\"," +
                "\"" + AssemblyVersion + "\"," +
                "\"" + FileVersion + "\"," +
                "\"" + ProductVersion + "\"," +
                "\"" + ImageRuntimeVersion + "\"," +
                "\"" + TargetFrameworkAttribute + "\"," +
                "\"" + PEKind + "\"," +
                "\"" + ImageFileMachine + "\"," +
                "\"" + ProcessorArchitecture + "\"," +
                "\"" + Error + "\"";
        }

        public static string[] CSVHeader
        {
            get
            {
                return new[] { "Full Path,FileName,AssemblyVersion,FileVersion,ProductVersion,ImageRuntimeVersion,TargetFrameworkAttribute,PortableExecutableKinds,ImageFileMachine,ProcessorArchitecture,Error" };
            }
        }

        public string AssemblyVersion { get; internal set; }
        public string FileVersion { get; internal set; }
        public string ProductVersion { get; internal set; }
    }
}
