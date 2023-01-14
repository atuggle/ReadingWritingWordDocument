using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        // Set the path to the Word document and the output.txt file
        string docPath = @"C:\Users\allen\code\scratch\ReadingWritingWordDocument\Template1.docx";
        string outputPath = @"C:\Users\allen\code\scratch\ReadingWritingWordDocument\output.txt";

        ReadWriteWordDocument(docPath, outputPath);
    }

    static void ReadWriteWordDocument(string wordDocument, string variableInputs)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(wordDocument, true))
        {
            var body = doc.MainDocumentPart.Document.Body;
            var paras = body.Elements<Paragraph>();

            // Read the output.txt file line by line
            using (var sr = new StreamReader(variableInputs))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();

                    // Split the line into key-value pairs
                    var keyValuePairs = line.Split('|');
                    var name = keyValuePairs[0];
                    var description = keyValuePairs[1];
                    var price = keyValuePairs[2];

                    // Replace all instances of the key surrounded by curly brackets in the Word document
                    ReplaceText(paras, "{name}", name);
                    ReplaceText(paras, "{description}", description);
                    ReplaceText(paras, "{price}", price);
                }
            }
        }
    }

    private static void ReplaceText(IEnumerable<Paragraph> paras, string textToReplace, string replacementText)
    {
        foreach (var para in paras)
        {
            foreach (var run in para.Elements<Run>())
            {
                foreach (var text in run.Elements<Text>())
                {
                    if (text.Text.Contains(textToReplace))
                    {
                        text.Text = text.Text.Replace(textToReplace, replacementText);
                    }
                }
            }
        }
    }
}
