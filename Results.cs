using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class Results : IDisposable {
      public ExcelPackage? book_ = null;
      public SummarySheet? Summary = null;
      public Dictionary<string, DocumentSheet> PageSheets = new Dictionary<string, DocumentSheet>();
      private string xlsxFilename_ = "searchResults.xlsx";
      private MatchedSheets? Matched;
      private UnmatchedSheets? Unmatched;

      public Results() {
         //xlsxFilename_ = $"searchResults-{DateTime.Now:yyyy-MM-dd HH-mm-ss}.xlsx";
         var fileInfo = new FileInfo(xlsxFilename_);
         if (fileInfo.Exists) {
            fileInfo.Delete();
            if (fileInfo.Exists) {
               throw new Exception($"Results file '{xlsxFilename_}' already exists and cannot be deleted");
               }
            }
         book_ = new ExcelPackage(fileInfo);

         // Create the Summary sheet that lists the keywords
         Summary = new SummarySheet(book_);

         // Create a worksheet that will list the documents according to whether they match any keyword
         Matched = new MatchedSheets(book_);
         Unmatched = new UnmatchedSheets(book_);
         }

      public void Dispose() {
         if (book_ != null) {
            book_.Save();
            book_.Dispose();
            book_ = null;
            }
         }

      public DocumentSheet AddPage(string pdfFilename, int pdfIndex, int pageNumber) {
         var pageSheet = new DocumentSheet(book_, pdfFilename, pdfIndex, pageNumber);
         PageSheets.Add(pdfFilename, pageSheet);
         return pageSheet;
         }

      internal void AddMatchedSheet(string pathName, string title, int pageCount) {
         Matched?.AddMatched(pathName, title, pageCount);
         }

      internal void AddUnmatchedSheet(string pathName, string title, int pageCount) {
         Unmatched?.addUnmatched(pathName, title, pageCount);
         }

      internal void Finish() {
         Summary?.Finish();
         Matched?.FormatColumns();
         Unmatched?.FormatColumns();

         foreach(var pageSheet in PageSheets) {
            pageSheet.Value.Finish();
            }
         }
      }
   }
