using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class MatchedSheets {
      private ExcelWorksheet? sheet_;
      private ExcelPackage book_;
      private int lastRow_ = -1;

      public MatchedSheets(ExcelPackage book) {
         book_ = book;
         sheet_ = book_.Workbook.Worksheets.Add("With matches");
         var cells = sheet_?.Cells;
         if (cells == null) {
            throw new Exception("Matched sheet has null Cells property");
            }
         cells[1, 1].Value = "The following files match at least one of the keywords:";
         cells[3, 1].Value = "Filename";
         cells[3, 2].Value = "Title";
         cells[3, 3].Value = "Pages";
         lastRow_ = 3;
         }

      public void addMatched(string pdfFilename, string title, int pageCount) {
         var cells = sheet_?.Cells;
         if (cells == null) {
            return;
            }
         cells[++lastRow_, 1].Value = pdfFilename;
         cells[lastRow_, 2].Value = title;
         cells[lastRow_, 3].Value = pageCount;
         }

      internal void FormatColumns() {
         var cells = sheet_?.Cells;
         if (cells == null) {
            return;
            }
         cells[1, 1].EntireColumn.AutoFit();
         cells[1, 2].EntireColumn.AutoFit();
         }
      }
   }
