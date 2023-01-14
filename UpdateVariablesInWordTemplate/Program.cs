using System;
using System.IO;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using UpdateVariablesInWordTemplate;

class Program
{
    static void Main(string[] args)
    {
        // TODO: pull the below two variables in from args
        // Set the path to the Word document and the output.txt file
        var templateFile = @"C:\Users\allen\code\scratch\ReadingWritingWordDocument\Template1.docx";
        var saleItemFile = @"C:\Users\allen\code\scratch\ReadingWritingWordDocument\output.txt";

        try
        {
            ReplaceVariablesInTemplate(templateFile, saleItemFile);
        } 
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

    }

    private static void ReplaceVariablesInTemplate(string templateFile, string saleItemFile)
    {
        var baseFolder = Path.GetDirectoryName(templateFile);

        // Read the output.txt file line by line
        using (var sr = new StreamReader(saleItemFile))
        {
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();

                // Split the line into key-value pairs
                var keyValuePairs = line.Split('|');
                var name = keyValuePairs[0];
                var description = keyValuePairs[1];
                var price = keyValuePairs[2];

                var saleItem = new SaleItem() { Name = name, Item = description, Price = price };
                SearchAndReplace(baseFolder, templateFile, saleItem);
            }
        }
    }

    // To search and replace content in a document part.
    private static void SearchAndReplace(string baseFolder, string templateFile, SaleItem saleItem)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateFile, false))
        {
            string docText = null;
            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            // Given this is a simple find/replace we do not need to load formal Open XML elements like Paragraphs, Runs, and Texts
            docText = docText.Replace("{name}", saleItem.Name);
            docText = docText.Replace("{description}", saleItem.Item);
            docText = docText.Replace("{price}", saleItem.Price);

            // Save new document in an output folder with the name of the saleitem.name
            var outputFolder = $@"{baseFolder}{Path.DirectorySeparatorChar}output";
            EnsureFolderExists(outputFolder);

            var newFileName = $@"{outputFolder}{Path.DirectorySeparatorChar}{saleItem.Name}.docx";
            RemoveFileIfExists(newFileName);
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

    private static void RemoveFileIfExists(string newFileName)
    {
        if (File.Exists(newFileName))
        {
            File.Delete(newFileName);
        }
    }

    private static void EnsureFolderExists(string path)
    {
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }
    }
}
