using System.Numerics;
using System.Reflection.Metadata.Ecma335;
using System.Text.RegularExpressions;
using System.Threading.Channels;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.Graphics.Operations.TextState;
using static System.Collections.Specialized.BitVector32;

namespace PdfSearch {
   internal class Program {
      static int Main(string[] args) {
         string folderPath = @"D:\root\WhitleyShaw\CAWS\Battery\WG\Documents\BESS PEIR";
         string[] spinner = new string[] { "|", "/", "-", "\\" };
         var spinCount = 0;
         var documents = new Dictionary<int, DocumentFile>();

         List<string> rawKeywords;
         try {
            rawKeywords = File.ReadAllLines("keywords.txt").ToList();
            }
         catch (Exception ex) {
            Console.WriteLine($"Could not read 'keywords.txt' from folder "
               + $"'{Directory.GetCurrentDirectory()}': {ex.Message}");
            throw;
            }
         if (!File.Exists("keywords.txt")) {
            return 2;
            }

         var keywords = rawKeywords.ConvertAll(kw => Regex.Escape(kw));
         var finder = new Regex(string.Join("|", keywords), RegexOptions.IgnoreCase);

         try {
            string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf");
            var fileCount = pdfFiles.Count();

            Console.WriteLine($"Scanning {fileCount} {pluralled("file", fileCount)} in folder {folderPath}:");
            foreach (string pdfFilename in pdfFiles) {
               using (var docFile = new DocumentFile()) {
                  docFile.Pathname = pdfFilename;
                  if (pdfFilename.Contains("Ch14")) {
                     docFile.Pathname = pdfFilename;
                     }
                  if (docFile.Open()) {
                     Console.Write($"\r\u001b[K\r{spinner[spinCount++ % 4]} {Path.GetFileName(pdfFilename)}"
                        + $", with {docFile.PageCount} {pluralled("page", docFile.PageCount)} ");

                     documents.Add(docFile.Id, docFile);
                     docFile.SearchPages(pdfFilename, keywords, finder);
                     }
                  else {
                     Console.WriteLine($"\r\u001b[K\rfailed to open PDF file '{pdfFilename}'");
                     }
                  }
               }
            return 0;
            }
         catch (Exception ex) {
            Console.WriteLine($"\r\u001b[K\rexception: {ex.Message}\u001b");
            return 1;
            }
         }

      internal static string pluralled(string word, int count) {
         return (count == 1) ? word : word + "s";
         }
      }
   }
