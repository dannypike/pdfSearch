// Object hierarchy:
//    DocumentFiles contain MatchingPages, which contain MatchingSentences indexed by keyword ids
//    keywords are back-linked to the sentences where they were be found

// PdfDocument is a class from the PdfPig library, which is a PDF document parser
// DocumentFile is our class that wraps the PdfDocument and provides search capabilities
// MatchingPage is a class that wraps a PdfPage and provides a place to store the sentences where keywords were found
// MatchingSentences is a class that stores the sentences where a keyword was found
// Keyword is a class that stores the text of a keyword and how it should be compared
// Program is the main class that drives
//    - the search for keywords in PDF files
//    - the display of the search progress, results and errors

using System.IO.Hashing;
using System;
using System.Runtime.Intrinsics.Arm;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using System.Text;
using System.Reflection.PortableExecutable;
using System.Diagnostics;

namespace PdfSearch {
   internal class DocumentFile : IDisposable {
      private static int nextId_ = 0;

      const int TruncateLength = 120;
      const string DefaultTitle = "";
      private Regex regexCleanForConsole = new Regex(@"[^\x20-\x7E]");
      private Results results_;  // The spreadsheet for output

      public int Id { get; set; } = ++nextId_;
      public string Pathname { get; set; } = "";

      private PdfDocument? pdfFile_ = null;

      public DocumentFile(Results results) {
         results_ = results;
         }

      public int PageCount {
         get {
            return pdfFile_?.NumberOfPages ?? 0;
            }
         }

      public int NumberOfPages { get; private set; }

      internal bool Open() {
         try {
            pdfFile_ = PdfDocument.Open(Pathname);
            return true;
            }
         catch (Exception ex) {
            Logger.WriteLine($"failed to open PDF file '{Pathname}'"
               + $", error was: {ex.Message}");
            return false;
            }
         }

      public void Dispose() {
         pdfFile_?.Dispose();
         }

      public bool SearchPages(string pathName, int pdfIndex, List<string> userKeywords
            , List<Regex?> regexKeywords, Regex finder, Results results) {

         string title = DefaultTitle;
         DocumentSheet? documentSheet = null;
         SummarySheet? summary = results.Summary;
         var pdfPageNumber = 0;   // The PDF page number (not the same as the Lime Down Index page number)
         NumberOfPages = pdfFile_?.NumberOfPages ?? 0;
         Program.TotalPages += NumberOfPages;
         while (++pdfPageNumber <= NumberOfPages) {
            var pdfPage = pdfFile_?.GetPage(pdfPageNumber);
            if (pdfPage == null) {
               Logger.WriteLine($"Failed to get page {pdfPageNumber} from PDF file '{pathName}'", false);
               break;
               }

            // Extract the blocks of text (cf. paragraphs) from the document
            var pdfLetters = pdfPage.Letters;
            var pdfWords = NearestNeighbourWordExtractor.Instance.GetWords(pdfLetters);
            var pdfUnorderedBlocks = DocstrumBoundingBoxes.Instance.GetBlocks(pdfWords);
            var pdfBlocks = UnsupervisedReadingOrderDetector.Instance.Get(pdfUnorderedBlocks).ToList();

            // The page number may be roman numerals or other non-numeric values and it can be anywhere on the page
            // (though typically near the beginning of the blocks).
            var pageNumber = new PageNumber() { PdfPageNumber = pdfPageNumber };

            // Look for keywords in each block
            var pdfBlockIndex = 0;
            var numberOfBlocks = pdfBlocks.Count();
            foreach (var pdfBlock in pdfBlocks) {
               // Is this the page number block?
               if (pageNumber.LdsPageNumber.Length == 0 && pdfBlock.Text.StartsWith("Page ")) {
                  var nextWord = pdfBlock.Text[5..];
                  if (nextWord.IndexOf(" ") < 0) {
                     // There is only one more word in the block, so it probably is what we want for the page number
                     if (pdfBlock.Text.Length > 5) {
                        pageNumber.LdsPageNumber = pdfBlock.Text[5..];
                        }
                     }
                  continue;
                  }

               var blockText = pdfBlock.Text;

               var searchText = blockText.Replace("\n", " ");

               // Is this the title block?
               if (title == DefaultTitle) {  // Have we already got the title?
                  int indexOfTitle = searchText.IndexOf("Volume ");
                  if (indexOfTitle >= 0) {
                     title = blockText[indexOfTitle..];
                     title = title.Replace("\n", " ").Split(".").FirstOrDefault() ?? title;
                     documentSheet?.SetTitle(title);
                     continue;
                     }
                  }

               // Look for keywords in this block
               string reportText;
               var result = finder.Matches(searchText);
               if (result.Count > 0) {
                  var reportLength = Math.Min(searchText.Length, TruncateLength);
                  if (0 < reportLength) {
                     if (reportLength >= searchText.Length) {
                        // It's a short sentence, so we can report it all
                        reportText = searchText;
                        }
                     else {
                        // Truncate the sentence to make it fit on a line
                        reportText = searchText.Substring(0, reportLength - 3) + "...";
                        }

                     // If this is the first match, then output the title
                     if (documentSheet == null) {
                        documentSheet = results.AddPage(pathName, pdfIndex, NumberOfPages);
                        Logger.WriteLine(title);
                        }

                     // And log the words that matched one of the keyword definitions
                     var matchingKeywords = result.Select(rr => rr.Value)
                        .Distinct(StringComparer.CurrentCultureIgnoreCase);

                     // Also record how many times each keyword matches any page in any document
                     var regexCount = regexKeywords.Count;
                     foreach (var keyword in matchingKeywords) {
                        for (var ii = 0; ii < regexCount; ++ii) {
                           var userKeyword = userKeywords[ii];
                           var isMatch = false;
                           var matchedWord = "";
                           if (userKeyword.StartsWith("/") && userKeyword.EndsWith("/")) {
                              // It's a regex definition string, so we use the supplied Regex
                              var regexKeyword = regexKeywords[ii];
                              isMatch = regexKeyword?.IsMatch(keyword) ?? false;
                              matchedWord = keyword;
                              }
                           else {
                              // Use simple string comparison for non-regex keywords
                              isMatch = userKeyword.Equals(keyword, StringComparison.CurrentCultureIgnoreCase);
                              matchedWord = keyword;
                              }
                           if (isMatch) {
                              summary?.IncKeyword(userKeyword, matchedWord);
                              }
                           }
                        }

                     // The console does not like non-ANSI codes
                     var consoleText = regexCleanForConsole.Replace(reportText, "\xa4");

                     Logger.WriteLine($"Page {pageNumber}: '{consoleText}' matches: "
                     + $"\"{string.Join("\", \"", matchingKeywords)}\"");

                     // Identify the block and include a CRC, in case it changes in the future
                     var crc32 = new System.IO.Hashing.Crc32();
                     crc32.Append(Encoding.UTF8.GetBytes(blockText));
                     var ll = crc32.HashLengthInBytes;
                     Debug.Assert(ll == 4);
                     
                     var blockId = $"{pdfPageNumber}/{BitConverter.ToUInt32(crc32.GetHashAndReset(), 0):X8}/{pdfBlockIndex}/{Program.Version}";
                     documentSheet?.AddKeywords(pageNumber, reportText, blockId, matchingKeywords);
                     }

                  if (summary != null) {
                     ++summary.TotalMatchingPages;
                     }
                  }
               ++pdfBlockIndex;
               }
            }

         if (0 == NumberOfPages) {
            Logger.WriteLine($"No pages in the PDF file '{pathName}'");
            }

         if (documentSheet != null) {
            documentSheet.FormatColumns();
            results.AddMatchedSheet(Path.GetFileName(pathName), title, NumberOfPages);

            // Document separator
            Logger.WriteLine("+++");
            return true;
            }
         results.AddUnmatchedSheet(Path.GetFileName(pathName), title, NumberOfPages);
         return false;
         }
      }
   }
