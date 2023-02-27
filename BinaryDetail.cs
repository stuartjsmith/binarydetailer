using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace BinaryDetailer
{
    [Serializable]
    public class BinaryDetail
    {

        public BinaryDetail(FileInfo fileInfo)
        {
            FileInfo = fileInfo;
        }

        public string Error { get; internal set; }
        public FileInfo FileInfo { get; internal set; }
        public ImageFileMachine ImageFileMachine { get; internal set; }
        public string ImageRuntimeVersion { get; internal set; }
        public PortableExecutableKinds PEKind { get; internal set; }
        public ProcessorArchitecture ProcessorArchitecture { get; internal set; }
        public string TargetFrameworkAttribute { get; internal set; }
        public string AssemblyVersion { get; internal set; }

        public string FileVersion => FileVersionInfo.GetVersionInfo(FileInfo.FullName).FileVersion;

        public string ProductVersion => FileVersionInfo.GetVersionInfo(FileInfo.FullName).ProductVersion;

        public string AssemblyCompanyAttribute => FileVersionInfo.GetVersionInfo(FileInfo.FullName).CompanyName;
        public string AssemblyCopyrightAttribute => FileVersionInfo.GetVersionInfo(FileInfo.FullName).LegalCopyright;


        public static string[] CSVHeader
        {
            get
            {
                return new[]
                {
                    "Full Path,FileName,AssemblyCompanyAttribute,AssemblyCopyrightAttribute,AssemblyVersion,FileVersion,ProductVersion,ImageRuntimeVersion,TargetFrameworkAttribute,PortableExecutableKinds,ImageFileMachine,ProcessorArchitecture,Error"
                };
            }
        }


        public string ToCsv()
        {
            return
                "\"" + FileInfo.FullName + "\"," +
                "\"" + FileInfo.Name + "\"," +
                "\"" + AssemblyCompanyAttribute + "\"," +
                "\"" + AssemblyCopyrightAttribute + "\"," +
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
    }
}