using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
      private int maxColumn_ = 0;
      private Dictionary<string, MatchedKeyword> keywordPages_
         = new Dictionary<string, MatchedKeyword>(StringComparer.CurrentCultureIgnoreCase);
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

         var lastRow = 0;

         cells[++lastRow, 1].Value = "Created:";
         cells[lastRow, 2].Value = $"{Program.Timestamp:dd MMM yyyy HH:mm:ss}";

         ++lastRow;

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
         cells[lastRow, 5].Value = "matches found";

         maxColumn_ = 5;
         keywordfirstRow_ = lastRow + 1;
         foreach (var kw in rawKeywords) {
            cells[++lastRow, 2].Value = kwId++;
            cells[lastRow, 3].Value = kw;

            if (keywordPages_.TryGetValue(kw, out var matchedKeyword)) {

               var column = 4;
               cells[lastRow, column].Value = matchedKeyword.Count;
               foreach (var matchWord in matchedKeyword.Matches) {
                  cells[lastRow, ++column].Value = matchWord;
                  }
               maxColumn_ = Math.Max(maxColumn_, column);
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
            cells[pageCountRow_, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
         if (matchingPageCountRow_ > 0) {
            cells[matchingPageCountRow_, 2].Value = totalMatchingPages_;
            }

         for (var row = keywordfirstRow_; row <= keywordLastRow_; ++row) {
            var kw = cells[row, 3].Text;
            if (keywordPages_.TryGetValue(kw, out var matchedKeyword)) {
               cells[row, 4].Value = matchedKeyword.Count;

               var columnIndex = 5;
               foreach (var mkw in matchedKeyword.Matches) {
                  cells[row, columnIndex++].Value = mkw;
                  }
               }
            }

         for (var columnIndex = 1; columnIndex <= maxColumn_; ++columnIndex) {
            cells[1, columnIndex].EntireColumn.AutoFit();
            }

         // Add the copyright after resizing the other columns
         cells[1, 5].Value = "Copyright (c) 2025 Community Action: Whitley and Shaw. All rights reserved.";
         cells[2, 5].Value = "This document is CONFIDENTIAL and MUST NOT be shown to any third parties, without the express permission from both the CAWS Chairman and CAWS Secretary.";
         }

      internal void IncKeyword(string userKeyword, string matchedWord) {
         if (!keywordPages_.TryGetValue(userKeyword, out var matchedKeyword)) {
            matchedKeyword = new MatchedKeyword();
            keywordPages_.Add(userKeyword, matchedKeyword);
            }
         matchedKeyword.AddCount(matchedWord);
         }
      }
   }
