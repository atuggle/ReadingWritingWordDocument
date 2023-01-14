using System;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using UpdateVariablesInWordTemplate;

class Program
{
    static void Main(string[] args)
    {
        // Set the path to the Word document and the output.txt file
        string filePath = @"C:\Users\allen\code\scratch\ReadingWritingWordDocument";
        string templateName = @"Template1.docx";
        //string outputPath = @"output.txt";

        var saleItem = new SaleItem() { Name = "John", Item = "Book", Price = "$15.00" };
        SearchAndReplace(filePath, templateName, saleItem);
    }

    // To search and replace content in a document part.
    private static void SearchAndReplace(string filePath, string templateName, SaleItem saleItem)
    {
        var templateFileName = $@"{filePath}{Path.DirectorySeparatorChar}{templateName}";
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateFileName, true))
        {
            string docText = null;
            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            Regex regexText = new Regex("{name}");
            docText = regexText.Replace(docText, saleItem.Name);

            regexText = new Regex("{description}");
            docText = regexText.Replace(docText, saleItem.Item);

            regexText = new Regex("{price}");
            docText = regexText.Replace(docText, saleItem.Price);

            var newFileName = $@"{filePath}{Path.DirectorySeparatorChar}{saleItem.Name}.docx";
            using (WordprocessingDocument newWordDoc = WordprocessingDocument.Create(newFileName, WordprocessingDocumentType.Document))
            {
                // Add the main document part to the new document
                newWordDoc.AddMainDocumentPart();

                using (StreamWriter sw = new StreamWriter(newWordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }

        }
    }
}
