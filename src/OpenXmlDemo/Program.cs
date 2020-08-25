using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using OpenXmlPowerTools;

namespace OpenXmlDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string sourceFile = Path.Combine("D:\\Word10.docx");
            string destinationFile = Path.Combine($"D:\\Word10-{Guid.NewGuid()}.docx");
            // Create a copy of the template file
            //File.Copy(sourceFile, destinationFile, true);
            //CreateTable(fileName);
            //SearchAndReplace(destinationFile);

            //File.WriteAllBytes(destinationFile, Replace4(File.ReadAllBytes(sourceFile)));
            File.WriteAllBytes(destinationFile, Replace5(File.ReadAllBytes(sourceFile)));


            Console.WriteLine("Terminó!");
        }

        // Insert a table into a word processing document.
        public static void CreateTable(string fileName)
        {
            // Use the file name and path passed in as an argument 
            // to open an existing Word 2007 document.

            using (WordprocessingDocument doc
                = WordprocessingDocument.Open(fileName, true))
            {
                // Create an empty table.
                Table table = new Table();

                // Create a TableProperties object and specify its border information.
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Thick),
                            Size = 24
                        }
                    )
                );

                // Append the TableProperties object to the empty table.
                table.AppendChild<TableProperties>(tblProp);

                // Create a row.
                TableRow tr = new TableRow();

                // Create a cell.
                TableCell tc1 = new TableCell();

                // Specify the width property of the table cell.
                tc1.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

                // Specify the table cell content.
                tc1.Append(new Paragraph(new Run(new Text("some text"))));

                // Append the table cell to the table row.
                tr.Append(tc1);

                // Create a second table cell by copying the OuterXml value of the first table cell.
                TableCell tc2 = new TableCell(tc1.OuterXml);

                // Append the table cell to the table row.
                tr.Append(tc2);

                // Append the table row to the table.
                table.Append(tr);

                // Append the table to the document.
                doc.MainDocumentPart.Document.Body.Append(table);
            }
        }

        //https://docs.microsoft.com/en-us/office/open-xml/how-to-search-and-replace-text-in-a-document-part
        // To search and replace content in a document part.
        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(@"@AQUI_SALUDO");
                docText = regexText.Replace(docText, "Hi Everyone!");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        //https://stackoverflow.com/a/18339407/2166103
        public static void Replace1(string document)
        {
            using (WordprocessingDocument wordDoc =
                      WordprocessingDocument.Open(@"yourpath\testdocument.docx", true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var paras = body.Elements<Paragraph>();

                foreach (var para in paras)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text.Contains("@SALUDO"))
                            {
                                text.Text = text.Text.Replace("@SALUDO", "Hola");
                            }
                        }
                    }
                }
            }
        }

        public static void Replace2(string document)
        {
            using (WordprocessingDocument wordDoc =
                    WordprocessingDocument.Open(@"yourpath\testdocument.docx", true))
            {
                var document_ = wordDoc.MainDocumentPart.Document;

                foreach (var text in document_.Descendants<Text>()) // <<< Here
                {
                    if (text.Text.Contains("@SALUDO"))
                    {
                        text.Text = text.Text.Replace("@SALUDO", "Hola");
                    }
                }
            }
        }
        //Replace que retorna un array de bytes
        public static byte[] Replace3(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(@"@AQUI_SALUDO");
                docText = regexText.Replace(docText, "Hi Everyone!");

                using (MemoryStream ms = new MemoryStream())
                {
                    wordDoc.MainDocumentPart.GetStream(FileMode.Create).CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }
        //Replace que retorna un array de bytes
        public static byte[] Replace4(byte[] arrayBytes)
        {
            using (MemoryStream stream = new MemoryStream(arrayBytes, true))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var document_ = wordDoc.MainDocumentPart.Document;

                foreach (var text in document_.Descendants<Text>()) // <<< Here
                {
                    if (text.Text.Contains("@AQUI_SALUDO"))
                    {
                        Regex regexText = new Regex(@"@AQUI_SALUDO");
                        text.Text = regexText.Replace(text.Text, "Hi Everyone!");
                    }
                }
                wordDoc.Save();
                //1) Si modifica el stream, defrente hacer el 
                return stream.ToArray();
                //2) else
                using (MemoryStream ms = new MemoryStream())
                {
                    wordDoc.MainDocumentPart.GetStream(FileMode.Create).CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }

        //https://github.com/EricWhiteDev/Open-Xml-PowerTools/blob/vNext/OpenXmlPowerToolsExamples/OpenXmlRegex02/OpenXmlRegex02.cs
        public static byte[] Replace5(byte[] arrayBytes)
        {
            using (MemoryStream stream = new MemoryStream(arrayBytes, true))
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(stream, true))
                {
                    //var xDoc = wDoc.MainDocumentPart.GetXDocument();
                    var xDoc = OpenXmlPowerTools.PtOpenXmlExtensions.GetXDocument(wDoc.MainDocumentPart);

                    var content = xDoc.Descendants(OpenXmlPowerTools.W.p);
                    Regex regex = new Regex("@AQUI_SALUDO");
                    OpenXmlPowerTools.OpenXmlRegex.Replace(content, regex, "Hi Everyone!", null);

                    //wDoc.MainDocumentPart.PutXDocument();
                    OpenXmlPowerTools.PtOpenXmlExtensions.PutXDocument(wDoc.MainDocumentPart);
                }
                return stream.ToArray();
            }
        }
    }
}
