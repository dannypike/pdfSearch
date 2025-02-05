// Copyright (c) 2024 Danny Pike.

using OfficeOpenXml;
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

         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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

         var keywords = rawKeywords.ConvertAll(kw => kw.StartsWith("/") && kw.EndsWith("/")
            ? kw[1..^1] : Regex.Escape(kw));
         var finder = new Regex(string.Join("|", keywords), RegexOptions.IgnoreCase);
         var matcher = keywords.ConvertAll(kw => (kw.StartsWith("/") && kw.EndsWith("/"))
            ? new Regex(kw[1..^1], RegexOptions.IgnoreCase) : new Regex(Regex.Escape(kw), RegexOptions.IgnoreCase));

         var now = DateTime.Now;
         try {
            string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf");
            var fileCount = pdfFiles.Count();

            // Create an Excel spreadsheet to hold the search results
            using (var results = new Results()) {
               Console.WriteLine($"Scanning {fileCount} {pluralled("file", fileCount)} in folder {folderPath}:");
               SummarySheet? summary = results.Summary;
               summary?.addKeywords(rawKeywords);

               foreach (string pdfFilename in pdfFiles) {
                  using (var docFile = new DocumentFile(results)) {
                     docFile.Pathname = pdfFilename;
                     if (docFile.Open()) {
                        Console.Write($"\r\u001b[K\r{spinner[spinCount++ % 4]} {Path.GetFileName(pdfFilename)}"
                           + $", with {docFile.PageCount} {pluralled("page", docFile.PageCount)} ");

                        documents.Add(docFile.Id, docFile);
                        var matched = docFile.SearchPages(pdfFilename, matcher, finder, results);
                        if (summary != null) {
                           ++summary.TotalFiles;
                           summary.TotalPages += docFile.NumberOfPages;
                           if (matched) {
                              ++summary.TotalMatchingFiles;
                              }
                           }
                        }
                     }

                  }
               results.Finish();
               return 0;
               }
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
