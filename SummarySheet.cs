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

      public SummarySheet(ExcelPackage book) {
         book_ = book;
         sheet_ = book_.Workbook.Worksheets.Add("Summary");
         }
      
      public void addKeywords(IEnumerable<string> rawKeywords) {
         var cells = sheet_?.Cells;
         if (cells == null) {
            throw new Exception("addKeywords:summary sheet has null Cells property");
            }
         cells[1, 1].Value = "#";
         cells[1, 2].Value = "Keyword";
         //cells[1, 3].Value = "Count";
         //cells[1, 4].Value = "Documents";

         var row = 1;
         var kwId = 1;
         foreach (var kw in rawKeywords) {
            cells[++row, 1].Value = kwId++;
            cells[row, 2].Value = kw;
            }
         }
      }
   }
