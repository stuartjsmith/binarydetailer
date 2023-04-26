using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
// Use of this library requires an install of Microsoft office
using Microsoft.Office.Interop.Word;

namespace BinaryDetailer
{
    public class Export
    {
        readonly List<string> ExcludeNames = new List<string>();

        public Export(List<string> excludeNames)
        {
            ExcludeNames = excludeNames;
        }
        public void GroupBinary(List<BinaryDetail> binaryDetails, string xmlFilePath)
        {
            List<BinaryDetail> groupedStrings = new List<BinaryDetail>();
            XDocument xmlDoc = XDocument.Load(xmlFilePath);

            foreach (var binaryDetail in binaryDetails)
            {
                // Set an escape from the loop once match is found
                bool groupIdSet = false;

                // Loop through each "group" element
                foreach (XElement group in xmlDoc.Root.Elements("group"))
                {
                    if (groupIdSet == true)
                    {
                        break;
                    }

                    string groupName = group.Attribute("name").Value;

                    if (binaryDetail.AssemblyCompanyAttribute == groupName)
                    {
                        // Loop through each "dll" element in the matching group
                        foreach (XElement dll in group.Elements("dll"))
                        {
                            // Check if the "dllname" attribute matches a "FileInfo.Name"
                            string dllName = dll.Attribute("dllname")?.Value ?? dll.Value;

                            if (binaryDetail.FileInfo.Name == dllName)
                            {
                                binaryDetail.GroupId = groupName;
                                groupIdSet = true;
                                break;
                            }
                        }
                    }
                    else 
                    {
                        binaryDetail.GroupId = "Unknown";
                    }
                }
            }
        }

        public void CreateReport(List<BinaryDetail> binaryDetails)
        {
            string csvFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                Guid.NewGuid() + ".csv");
            File.Create(csvFileName).Close();
            File.AppendAllLines(csvFileName, BinaryDetail.CSVHeader);

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

                File.AppendAllLines(csvFileName, new[] { binaryDetail.ToCsv() });
            }

            Console.WriteLine("File created at " + csvFileName);
        }

        public void CreateWordDoc(List<BinaryDetail> binaryDetails)
        {
            try
            {
                string wordFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    Guid.NewGuid() + ".docx");

                //Create an instance for word app  
                Application winword = new Application
                {
                    ShowAnimation = false,
                    Visible = false
                };

                //Create a missing variable for missing value  
                object missing = Missing.Value;

                //Create a new document  
                Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                //Add header into the document  
                foreach (Section section in document.Sections)
                {
                    //Get the header range and add the header details.  
                    Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Binary Detailer";
                }

                //Add the footers into the document  
                foreach (Section wordSection in document.Sections)
                {
                    //Get the footer range and add the footer details.  
                    Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Github repository: https://github.com/stuartjsmith/binarydetailer";
                }

                //adding text to document  
                document.Content.SetRange(0, 0);
                string firstLine = "Binary Detailer is a method of outputting binary details such as net framework, " +
                    "64-bit compatibility, version etc. Given a directory name, it will iterate dll and exe files " +
                    "and output the results to a csv file.";

                document.Content.Text = firstLine + Environment.NewLine;

                //Add paragraph with Heading 1 style  
                Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Heading 1";
                para1.Range.set_Style(ref styleHeading1);
                para1.Range.Text = "Binary Details";
                para1.Range.InsertParagraphAfter();

                //Create a 5X5 table and insert some dummy record (+1 in count to allow for a column heading)
                Table firstTable = document.Tables.Add(para1.Range, binaryDetails.Count + 1, 4, ref missing, ref missing);

                // Define column headings
                WordColumnHeadings(firstTable);

                firstTable.Borders.Enable = 1;
                foreach (var binaryDetail in binaryDetails.Select((value, i) => new { i, value }))
                {

                    bool include = true;
                    foreach (string excludeName in ExcludeNames)
                    {
                        if ((binaryDetail.value.AssemblyCompanyAttribute != null && binaryDetail.value.AssemblyCompanyAttribute.ToLower().Contains(excludeName)) ||
                            (binaryDetail.value.AssemblyCopyrightAttribute != null && binaryDetail.value.AssemblyCopyrightAttribute.ToLower().Contains(excludeName)))
                        {
                            include = false;
                        }
                    }

                    if (include == false) continue;

                    // Passing the index + 2 as there is a heading column to overcome.
                    WordRowData(binaryDetail.i + 2, firstTable, binaryDetail.value);
                }

                //Save the document  
                object filename = wordFileName;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;

                Console.WriteLine("Word Document created at " + filename);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void WordColumnHeadings(Table firstTable)
        {
            // Loop through each cell in the row
            foreach (Cell cell in firstTable.Rows[1].Cells)
            {
                // Switch between the columns
                switch (cell.ColumnIndex)
                {
                    case 1:
                        cell.Range.Text = "Group ID";
                        break;
                    case 2:
                        cell.Range.Text = "File Name";
                        break;
                    case 3:
                        cell.Range.Text = "Assembly Version";
                        break;
                    case 4:
                        cell.Range.Text = "File Version";
                        break;
                }

                // Format properties goes here  
                cell.Range.Font.Bold = 1;
                cell.Range.Font.Name = "verdana";
                cell.Range.Font.Size = 10;

                // cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;

                // Center alignment for the Header cells  
                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        public void WordRowData(int index, Table firstTable, BinaryDetail binaryDetail)
        {
            // Loop through each cell in the row
            foreach (Cell cell in firstTable.Rows[index].Cells)
            {
                // Switch between the columns
                switch (cell.ColumnIndex)
                {
                    case 1:
                        cell.Range.Text = binaryDetail.GroupId;
                        break;
                    case 2:
                        cell.Range.Text = binaryDetail.FileInfo.Name;
                        break;
                    case 3:
                        cell.Range.Text = binaryDetail.AssemblyVersion;
                        break;
                    case 4:
                        cell.Range.Text = binaryDetail.FileVersion;
                        break;
                }
            }
        }
    }
}
