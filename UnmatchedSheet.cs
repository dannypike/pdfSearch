using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class UnmatchedSheet {
      private ExcelWorksheet? sheet_;
      private ExcelPackage book_;
      private int lastRow_ = -1;

      public UnmatchedSheet(ExcelPackage book) {
         book_ = book;
         sheet_ = book_.Workbook.Worksheets.Add("No matches");
         var cells = sheet_?.Cells;
         if (cells == null) {
            throw new Exception("Unmatched sheet has null Cells property");
            }
         cells[1, 1].Value = "The following files do not match any of the keywords:";
         cells[3, 1].Value = "Filename";
         cells[3, 2].Value = "Title";
         lastRow_ = 3;
         }

      public void addUnmatched(string pdfFilename, string title) {
         var cells = sheet_?.Cells;
         if (cells == null) {
            return;
            }
         cells[++lastRow_, 1].Value = pdfFilename;
         cells[lastRow_, 2].Value = title;
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
