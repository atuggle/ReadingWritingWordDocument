using System;
using System.IO;
using Microsoft.Office.Interop.Word;

class Program
{
    static void Main(string[] args)
    {
        // Set the path to the Word document and the output.txt file
        string docPath = @"C:\example.docx";
        string outputPath = @"C:\output.txt";

        // Create a new instance of the Word application
        Application app = new Application();

        // Open the Word document
        Document doc = app.Documents.Open(docPath);

        // Read in the output.txt file
        string[] lines = File.ReadAllLines(outputPath);

        // Iterate through each line in the output.txt file
        foreach (string line in lines)
        {
            // Split the line into a key-value pair
            string[] parts = line.Split('=');
            string key = parts[0].Trim();
            string value = parts[1].Trim();

            // Replace all instances of the key surrounded by curly brackets in the document with the value
            doc.Content.Find.ClearFormatting();
            doc.Content.Find.Text = "{" + key + "}";
            doc.Content.Find.Replacement.ClearFormatting();
            doc.Content.Find.Replacement.Text = value;
            doc.Content.Find.Execute(Replace: WdReplace.wdReplaceAll);
        }

        // Save the modified document
        doc.Save();

        // Close the document and the Word application
        doc.Close();
        app.Quit();
    }
}
