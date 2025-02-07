using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class SummarySheet {
      private ExcelWorksheet sheet_;
      private ExcelPackage book_;
      private int totalPages_ = 0;
      private int totalMatchingPages_ = 0;
      private int pageCountRow_;
      private int matchingPageCountRow_;
      private int totalFiles_ = 0;
      private int totalMatchingFiles_ = 0;
      private int matchingFileCountRow_;
      private int fileCountRow_;
      private Dictionary<string, int> keywordPages_
         = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
      private int keywordfirstRow_;
      private int keywordLastRow_;

      public SummarySheet(ExcelPackage book) {
         book_ = book;
         sheet_ = book_.Workbook.Worksheets.Add("Summary");
         }

      public int TotalMatchingPages {
         get => totalMatchingPages_; set => totalMatchingPages_ = value;
         }

      public int TotalPages {
         get => totalPages_;
         internal set => totalPages_ = value;
         }

      public int TotalMatchingFiles {
         get => totalMatchingFiles_;
         set => totalMatchingFiles_ = value;
         }

      public int TotalFiles {
         get => totalFiles_;
         set => totalFiles_ = value;
         }

      public void addKeywords(IEnumerable<string> rawKeywords) {
         var cells = sheet_?.Cells;
         if (cells == null) {
            throw new Exception("addKeywords:summary sheet has null Cells property");
            }
         //cells[1, 3].Value = "Count";
         //cells[1, 4].Value = "Documents";

         var lastRow = 0;

         cells[++lastRow, 1].Value = "# of documents:";
         fileCountRow_ = lastRow;
         cells[++lastRow, 1].Value = "# of documents with matches:";
         matchingFileCountRow_ = lastRow;
         cells[++lastRow, 1].Value = "# of pages:";
         pageCountRow_ = lastRow;
         cells[++lastRow, 1].Value = "# of pages with matches:";
         matchingPageCountRow_ = lastRow;
         ++lastRow;
         cells[++lastRow, 1].Value = "Keywords:";

         var kwId = 1;
         cells[++lastRow, 2].Value = "#";
         cells[lastRow, 3].Value = "Keyword";
         cells[lastRow, 4].Value = "# of pages";

         keywordfirstRow_ = lastRow + 1;
         foreach (var kw in rawKeywords) {
            cells[++lastRow, 2].Value = kwId++;
            cells[lastRow, 3].Value = kw;

            int count = 0;
            keywordPages_.TryGetValue(kw, out count);
            if (count > 0) {
               cells[lastRow, 4].Value = count;
               }
            }
         keywordLastRow_ = lastRow;
         }

      internal void Finish() {
         var cells = sheet_?.Cells;
         if (cells == null) {
            return;
            }
         if (fileCountRow_ > 0) {
            cells[fileCountRow_, 2].Value = totalFiles_;
            }
         if (matchingFileCountRow_ > 0) {
            cells[matchingFileCountRow_, 2].Value = totalMatchingFiles_;
            }
         if (pageCountRow_ > 0) {
            cells[pageCountRow_, 2].Value = totalPages_;
            }
         if (matchingPageCountRow_ > 0) {
            cells[matchingPageCountRow_, 2].Value = totalMatchingPages_;
            }

         for (var row = keywordfirstRow_; row <= keywordLastRow_; ++row) {
            var kw = cells[row, 3].Text;
            if (keywordPages_.TryGetValue(kw, out int count)) {
               cells[row, 4].Value = count;
               }
            }

         cells[1, 1].EntireColumn.AutoFit();
         cells[1, 2].EntireColumn.AutoFit();
         cells[1, 3].EntireColumn.AutoFit();
         }

      internal void IncKeyword(string userKeyword) {
         keywordPages_.TryGetValue(userKeyword, out int count);
         keywordPages_[userKeyword] = count + 1;
         }
      }
   }
