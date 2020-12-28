﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TRIS.FormFill.Lib;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string merged = @"D:\asthenis\DocxPaging\merged.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(merged, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;

                var firstSectionProp = mainPart.Document.Body.Elements<SectionProperties>().First();
                firstSectionProp.Append(new PageNumberType { Start = 1 });

                AppendFooter(mainPart, "Page ", "2");
                AppendSectionBreaks(mainPart);
                doc.Close();
            }

            /*
            string template = @"E:\ConsoleApp1\files\Arztbrief_Ackermann_Detlef_1932-08-13-copy.docx";
            
            //string template = @"D:\asthenis\DocxPaging\TwoPagedTemplate.docx";
            //string tempalte2 = @"E:\ConsoleApp1\files\Arztbrief_Ackermann_Detlef_1932-08-13.docx";

            string datafile = @"E:\ConsoleApp1\files\Arztbrief_Ackermann_Detlef_1932-08-13.csv";
            string outputfile = @"E:\ConsoleApp1\files\output.docx";

            using (FileStream fs = new FileStream(outputfile, FileMode.OpenOrCreate))
            {
                var ms = Merge(template, datafile, true);
                ms.Seek(0, SeekOrigin.Begin);
                ms.CopyTo(fs);
                fs.Close();
            }
            // run MS WORD
            System.Diagnostics.Process.Start(outputfile);
            */

            /*
            string outputfile2 = @"D:\asthenis\DocxPaging\GenFooter.docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Create(outputfile2, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                body.AppendChild(new Paragraph(new Run(new Text("Hallo there"))));

                AppendFooter(mainPart, "Page 1 of 1");

                doc.Close();
            }
            System.Diagnostics.Process.Start(outputfile2);
            */
        }

        private static void AppendSectionBreaks(MainDocumentPart mainPart)
        {
            if (mainPart == null)
            {
                return;
            }

            //string sectionBreakPara = "<w:p><w:pPr><w:sectPr><w:type w:val=\"nextPage\" /></w:sectPr></w:pPr></w:p>";
            //Paragraph p = new Paragraph(sectionBreakPara);
            //string FooterId = mainPart.GetIdOfPart(mainPart.FooterParts.FirstOrDefault());
            //SectionType SectionBreakType = new SectionType() { Val = SectionMarkValues.NextPage };

            
            Break[] breaks = mainPart.Document.Descendants<Break>().ToArray();
            
            if (breaks.Length > 0)
            {
                for (int i = 0; i <= breaks.Length - 1; i++)
                {
                    var p = new Paragraph(new SectionProperties(new PageNumberType { Start = 1 }, new SectionType() { Val = SectionMarkValues.NextPage }));
                    var firstParagraph = breaks[i].Parent.Parent.NextSibling();
                    firstParagraph.InsertBeforeSelf(p);
                }
            }
            mainPart.Document.Save();
        }

        private static string AppendFooter(MainDocumentPart mainPart, string footerText, string TemplatePages)
        {
            // 1. check mainPart
            if (mainPart == null) return null;

            // 2. remove all footers if present
            if (mainPart.FooterParts.Count() > 0)
            {
                mainPart.DeleteParts(mainPart.FooterParts);
            }

            // 3. create and add my own footerPart
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();

            // 4. create Footer in my footerPart
            footerPart.Footer = new Footer(new Paragraph(new Run(
                new Text() { Text = footerText, Space = SpaceProcessingModeValues.Preserve },
                new SimpleField() { Instruction = "PAGE" },
                new Text() { Text = " of ", Space = SpaceProcessingModeValues.Preserve },
                new Text() { Text = TemplatePages, Space = SpaceProcessingModeValues.Preserve }
            )));

            // 5.  create footerPart reference
            string footerPartId = mainPart.GetIdOfPart(footerPart);
            FooterReference FooterRef = new FooterReference() { Type = HeaderFooterValues.Default, Id = footerPartId };

            // 6. find SectionProperties of document to insert my footerPart reference
            SectionProperties sectionProp = mainPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProp == null)
            {
                sectionProp = new SectionProperties();                
                mainPart.Document.Body.Append(sectionProp);
            }
            sectionProp.InsertAt(FooterRef, 0);

            return footerPartId;

            //sectionProp.Append(new PageNumberType { Start = 1 });
            //SectionType SectionBreakType = new SectionType() { Val = SectionMarkValues.NextPage };
            //sectionProp.Append(SectionBreakType);
        }

        public static DataSet Loadcsv(string CSV)
        {
            if (!File.Exists(CSV)) throw new ArgumentNullException("Cannot find " + CSV);
            string line;
            MatchCollection columns;
            DataSet dataSet = new DataSet();
            char[] charArray = { ',' };
            DataTable dataTable = dataSet.Tables.Add(CSV);

            using (FileStream stream = new FileStream(CSV, FileMode.Open))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    line = reader.ReadLine();
                    var regex = new Regex("(?:^|,)(\"(?:[^\"]+|\"\")*\"|[^,]*)");
                    regex.Matches(line);

                    columns = regex.Matches(line);

                    int i = 0;
                    while (i < columns.Count)
                    {
                        dataTable.Columns.Add(columns[i].Value.Trim('"').Trim(',').Trim('"').Trim());
                        i++;
                    }

                    line = reader.ReadLine();
                    while (line != null)
                    {
                        columns = regex.Matches(line);
                        var dataRow = dataTable.NewRow();
                        for (int j = 0; j < columns.Count; j++)
                        {
                            dataRow[j] = columns[j].Value.Trim('"').Trim(',').Trim('"').Trim();
                        }
                        dataTable.Rows.Add(dataRow);
                        line = reader.ReadLine();
                    }
                    reader.Close();
                }
            }
            return dataSet;
        }

        public static Dictionary<string, string> CreateOneRoll(DataColumnCollection cols, DataRow row)
        {
            Dictionary<string, string> d = new Dictionary<string, string>();
            for (int i = 0; i < cols.Count - 1; i++)
            {
                d.Add(cols[i].Caption, row[cols[i].Caption].ToString());
            }
            return d;
        }

        public static Stream Merge(string TemplateFile, string CSVFile, bool InsertPageBreaksBetweenPass = true)
        {
            if (!File.Exists(TemplateFile))
                throw new FileNotFoundException("Cannot find file", TemplateFile);

            if (!File.Exists(CSVFile))
                throw new FileNotFoundException("Cannot find file", CSVFile);

            MemoryStream outStream = new MemoryStream();

            byte[] buf = File.ReadAllBytes(TemplateFile);

            outStream.Write(buf, 0, buf.Length);

            Paragraph PageBreakParagraph = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));

            using (WordprocessingDocument outDoc = WordprocessingDocument.Open(outStream, true))
            {
                XElement newBody = XElement.Parse(outDoc.MainDocumentPart.Document.Body.OuterXml);
                //wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text
                // remove mailmerge if exists
                DocumentSettingsPart settingsPart = outDoc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
                MailMerge mm = settingsPart.Settings.OfType<MailMerge>().FirstOrDefault();

                if (mm != null)
                {
                    settingsPart.Settings.RemoveChild(mm);
                    settingsPart.Settings.Save();
                }

                // Load CSV-Data
                DataSet csv = Loadcsv(CSVFile);
                // run every csv row
                int i = 0;
                string OneMergePages = "0";

                foreach (DataRow dtRow in csv.Tables[0].Rows)
                {
                    Dictionary<string, string> roll = CreateOneRoll(csv.Tables[0].Columns, dtRow);
                    XElement tempBody = FormFiller.GetWordReport(TemplateFile, null, roll, out OneMergePages);// as MemoryStream do begin          
                    //Body tempBody = FormFiller.GetWordReport(TemplateFile, null, roll, out OneMergePages);// as MemoryStream do begin          

                    if (i == 0)
                    {
                        //Body b = new Body(tempBody.ToString());
                        //b.AppendChild(new Paragraph(new Text($"1 of {pages}")));

                        //  https://stackoverflow.com/questions/11947301/restart-page-numbering-in-header-with-openxml-sdk-2-0
                        // 
                        // http://officeopenxml.com/WPSectionPgNumType.php
                        // should i insert SectionBreak after each roll?

                        // add section breaks by hands: https://edu.gcfglobal.org/en/word2016/page-numbers/1/


                        #region add section break (not used)
                        Paragraph paragraph232 = new Paragraph();
                        ParagraphProperties paragraphProperties220 = new ParagraphProperties();
                        SectionProperties sectionProperties1 = new SectionProperties();
                        SectionType sectionType1 = new SectionType() { Val = SectionMarkValues.NextPage };
                        sectionProperties1.Append(sectionType1);
                        paragraphProperties220.Append(sectionProperties1);
                        paragraph232.Append(paragraphProperties220);
                        #endregion

                        //tempBody.Append(paragraph232);
                        newBody.ReplaceNodes(tempBody.Elements());
                    }
                    else
                    {
                        if (InsertPageBreaksBetweenPass == true)
                        {
                            newBody.Add(XElement.Parse(PageBreakParagraph.OuterXml));
                            newBody.Add(tempBody.Elements());
                        }
                        else
                        {
                            newBody.Add(tempBody.Elements());
                        }
                    }
                    i++;
                }

                outDoc.MainDocumentPart.Document.Body.Remove();
                outDoc.MainDocumentPart.Document.Body = new Body(newBody.ToString());
                outDoc.MainDocumentPart.Document.Save();

                MainDocumentPart mainPart = outDoc.MainDocumentPart;

                AppendFooter(mainPart, "Page ", OneMergePages);
                mainPart.Document.Save();


            } // End Using
            return outStream;
        }


        // not used
        private static void GenerateFooterPart1Content(FooterPart footerPart1, int page, int pages)
        {
            Footer footer1 = new Footer();
            footer1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Paragraph paragraph1 = new Paragraph();
            //paragraph1.Append(new SimpleField() { Instruction = "PAGE" });
            paragraph1.Append(new Text($"{page} of {pages}"));
            //paragraph1.Append(new SimpleField() { Instruction = "NUMPAGES" });

            footer1.Append(paragraph1);
            footerPart1.Footer = footer1;
        }

        // not used
        private static Footer GeneratePageFooterPart(string FooterText)
        {
            var element =
                new Footer(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = "Footer" }),
                        new Run(
                            new Text(FooterText))));
            // *** Adaptation: This will output the page number dynamically ***
            //new SimpleField() { Instruction = "PAGE" })
            //);

            return element;
        }
    }
}
/*
                 FooterPart footerPart = mainPart.FooterParts.FirstOrDefault();
                if (footerPart == null)
                {
                    footerPart = mainPart.AddNewPart<FooterPart>();
                }

                //footerPart.Footer = GeneratePageFooterPart("Footer texzt");
                GenerateFooterPart1Content(footerPart);
                
                string footerPartRef = mainPart.GetIdOfPart(footerPart);
                FooterReference objFooterReference = new FooterReference()
                {
                    Type = HeaderFooterValues.Default,
                    Id = footerPartRef
                };

                IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers                    
                    section.RemoveAllChildren<FooterReference>();

                    // Create the new header and footer reference node                    
                    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartRef });
                }
                    //mainPart.Document.Body.Append(oSectionProperties);
 */

// add footer part
// https://stackoverflow.com/questions/11672991/add-header-and-footer-to-an-existing-empty-word-document-with-openxml-sdk-2-0
// https://stackoverflow.com/questions/38430658/how-to-dynamically-add-a-page-number-in-footer-in-microsoft-oxml-c-sharp
// https://woodsworkblog.wordpress.com/2012/08/06/add-header-and-footer-to-an-existing-word-document-with-openxml-sdk-2-0/