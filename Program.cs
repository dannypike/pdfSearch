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
      public static int TotalPages { get; internal set; }
      public static string Version { get => "1"; }
      public static DateTime Timestamp { get; internal set; } = DateTime.Now;

      static int Main(string[] args) {
         string folderPath = @"D:\root\WhitleyShaw\CAWS\Battery\WG\Documents\BESS PEIR";
         string[] spinner = new string[] { "|", "/", "-", "\\" };
         var spinCount = 0;
         var documents = new Dictionary<int, DocumentFile>();

         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

         List<string> rawKeywords;
         try {
            rawKeywords = File.ReadAllLines("keywords.txt")
               .Select(kw => kw.Trim())
               .ToList();
            rawKeywords.RemoveAll(kw => kw.StartsWith("#") || (kw == ""));
            }
         catch (Exception ex) {
            Console.WriteLine($"Could not read 'keywords.txt' from folder "
               + $"'{Directory.GetCurrentDirectory()}': {ex.Message}");
            throw;
            }
         if (!File.Exists("keywords.txt")) {
            return 2;
            }

         string definition;
         List<string> keywords = new List<string>();
         List<Regex?> individualRegexes = new List<Regex?>();
         foreach (var kw in rawKeywords) {
            if (kw.StartsWith("/") && kw.EndsWith("/")) {
               definition = $"({kw[1..^1]})";
               individualRegexes.Add(new Regex(definition, RegexOptions.IgnoreCase));
               }
            else {
               definition = kw;
               individualRegexes.Add(null);  // Use a text comparison, not a Regex
               }
            keywords.Add(definition);
            }

         var quickFinder = new Regex(string.Join("|", keywords), RegexOptions.IgnoreCase);

         var now = DateTime.Now;
         var consoleTitle = OperatingSystem.IsWindows() ? Console.Title : "PdfSearch";
         try {
            string[] pdfFiles = [];
            if (File.Exists("filenames.txt")) {
               pdfFiles = File.ReadAllLines("filenames.txt")
                  .Select(tt => tt.Trim())
                  .Where(tt => !tt.StartsWith("# "))
                  .ToArray();
               }
            if (0 == pdfFiles.Length) {
               pdfFiles = Directory.GetFiles(folderPath, "*.pdf");
               }
            var fileCount = pdfFiles.Count();

            // Create an Excel spreadsheet to hold the search results
            using (var results = new Results()) {
               Logger.WriteLine($"Scanning {fileCount} {pluralled("file", fileCount)} in folder {folderPath}:");
               SummarySheet? summary = results.Summary;
               summary?.addKeywords(rawKeywords);

               int pdfIndex = 0;
               foreach (string pdfFilename in pdfFiles) {
                  using (var docFile = new DocumentFile(results)) {
                     docFile.Pathname = pdfFilename;
                     if (docFile.Open()) {
                        var pageCount = docFile.PageCount;

                        Console.Write($"\r\u001b[K\r{spinner[spinCount++ % 4]} {Path.GetFileName(pdfFilename)}"
                           + $", with {pageCount} {pluralled("page", pageCount)} ");

                        documents.Add(docFile.Id, docFile);

                        if (OperatingSystem.IsWindows()) {
                           var trimmedName = Path.GetFileNameWithoutExtension(pdfFilename);
                           var trimIndex = trimmedName.IndexOf("_V");
                           if (trimIndex > 0) {
                              trimmedName = trimmedName[(trimIndex + 1)..];
                              }
                           if (OperatingSystem.IsWindows()) {
                              Console.Title = $" ({pdfIndex};{pageCount})-{trimmedName}";
                              }
                           }
                        var matched = docFile.SearchPages(pdfFilename, ++pdfIndex, rawKeywords
                           , individualRegexes, quickFinder, results);
                        if (summary != null) {
                           ++summary.TotalFiles;
                           summary.TotalPages += pageCount;
                           if (matched) {
                              ++summary.TotalMatchingFiles;
                              }
                           }
                        }
                     }

                  }
               results.Finish();
               Logger.WriteLine($"Scanned {TotalPages} pages over {fileCount} {pluralled("document", fileCount)}");
               return 0;
               }
            }
         catch (Exception ex) {
            Logger.WriteLine($"exception: {ex.Message}");
            return 1;
            }
         finally {
            if (OperatingSystem.IsWindows()) {
               Console.Title = consoleTitle;
               }
            }
         }

      internal static string pluralled(string word, int count) {
         return (count == 1) ? word : word + "s";
         }
      }
   }
