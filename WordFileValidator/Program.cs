using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordFileValidator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = string.Empty;

            Thread t = new Thread((ThreadStart)(() => {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = "c:\\",
                    Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*",
                    FilterIndex = 2,
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                }
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Error, file not found at path: " + filePath);
                Console.ReadKey();

                return;
            }

            ValidateWordDocument(filePath);

            Console.WriteLine("The file is valid so far.");
            Console.WriteLine("Inserting some text into the body that would cause Schema error");
            Console.WriteLine("------ Press any key to continue ------");
            Console.ReadKey();

            ValidateCorruptedWordDocument(filePath);

            Console.WriteLine("All done! Press any key.");
            Console.ReadKey();
        }

        public static void ValidateWordDocument(string filePath)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, true))
            {
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();

                    int count = 0;

                    foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument))
                    {
                        count++;

                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine("count={0}", count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                wordprocessingDocument.Close();
            }
        }

        public static void ValidateCorruptedWordDocument(string filePath)
        {
            // Insert some text into the body, this would cause Schema Error
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, true))
            {
                // Insert some text into the body, this would cause Schema Error
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                Run run = new Run(new Text("some text"));

                body.Append(run);

                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();

                    int count = 0;

                    foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument))
                    {
                        count++;

                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine("count={0}", count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
