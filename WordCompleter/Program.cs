using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using SmartFormat;

namespace WordCompleter
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("usage:");
                Console.WriteLine("template.docx data.json result.docx");

                return;
            }

            var templatePath = args[0];
            var dataPath = args[1];
            var outputPath = args[2];

            var dataText = File.ReadAllText(dataPath);
            var data = JsonConvert.DeserializeObject<dynamic>(dataText);

            using var stream = File.OpenRead(templatePath);
            using var streamCopy = new MemoryStream();
            stream.CopyTo(streamCopy);
            
            using var document = WordprocessingDocument.Open(streamCopy, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            ReplacePlaceholders(body, data);
            
            document.SaveAs(outputPath);
        }

        private static void ReplacePlaceholders(Body body, dynamic data)
        {
            foreach (var text in body.Descendants<Text>())
            {
                if (text.Text.Contains("{"))
                {
                    text.Text = Smart.Format(text.Text, data);
                }
            }
        }
    }
}
