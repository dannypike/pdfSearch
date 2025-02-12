using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class DocumentSheet {
      private ExcelPackage? book_;
      private ExcelWorksheet? sheet_;
      private ExcelRange? range_;
      private string pdfFilename_;
      private int numberOfPages_;
      private int firstRow_;
      private int nextRow_;
      private Dictionary<int, ExcelRange> keywordRows_ = new Dictionary<int, ExcelRange>();
      private int titleRow_;
      private int maxMatchingColumn_ = 4;
      private string titleText_ = "";
      const string REMOVE_FILENAME_COMMON_PART1 = "EN010168 LDSP PEIR ";
      const int COLUMN_WIDTH_CONTEXT = 115;
      private Dictionary<int, PageNumber> toCheck_ = new Dictionary<int, PageNumber>();
      private int pageToReadRow_;

      internal int DocumentIndex { get; private set; }

      public DocumentSheet(ExcelPackage? book, string pdfFilename, int pdfIndex, int numberOfPages) {
         book_ = book;
         pdfFilename_ = pdfFilename;
         numberOfPages_ = numberOfPages;
         DocumentIndex = pdfIndex;

         var fileInfo = new FileInfo(pdfFilename_);
         var namePrefix = Path.GetFileNameWithoutExtension(fileInfo.Name);
         namePrefix = namePrefix
            .Replace("_", " ")
            .Replace(REMOVE_FILENAME_COMMON_PART1, "")
            ;
         var xlSheets = book_?.Workbook.Worksheets;
         if (xlSheets == null) {
            throw new Exception("failed to add document sheet: there is no workbook");
            }
         if (xlSheets.Any(s => s.Name == namePrefix)) {
            throw new Exception($"failed to add document sheet: sheet '{namePrefix}' already exists");
            }
         sheet_ = xlSheets.Add(namePrefix);
         sheet_.Column(3).Width = COLUMN_WIDTH_CONTEXT;

         range_ = sheet_?.Cells;
         if (range_ == null) {
            throw new Exception($"document sheet '{namePrefix}' has null Cells property");
            }

         // Summary details
         nextRow_ = range_.Start.Row;
         range_[nextRow_, 1].Value = "Document id";
         range_[nextRow_, 2].Value = $"{Program.Version}.{DocumentIndex}";
         ++nextRow_;

         range_[++nextRow_, 1].Value = "Document details";

         range_[++nextRow_, 1].Value = "PDF file";
         range_[nextRow_, 3].Value = Path.GetFileName(pdfFilename_);

         range_[++nextRow_, 1].Value = "Title";
         range_[nextRow_, 3].Value = "not found";
         titleRow_ = nextRow_;

         range_[++nextRow_, 1].Value = "Total Pages";
         range_[nextRow_, 2].Value = numberOfPages;

         range_[++nextRow_, 1].Value = "Pages to read";

         pageToReadRow_ = nextRow_;

         ++nextRow_;
         ++nextRow_;
         range_[++nextRow_, 1].Value = "Matches found";
         ++nextRow_;

         // List of paragraphs where keywords were found
         firstRow_ = ++nextRow_;
         range_[nextRow_, 1].Value = "ID (ignore this)";
         range_[nextRow_, 2].Value = "Page";
         range_[nextRow_, 3].Value = "Context";
         range_[nextRow_, 4].Value = "Keywords";
         }

      internal void SetTitle(string title) {
         titleText_ = title;
         if (range_ != null) {
            range_[titleRow_, 3].Value = titleText_;
            }
         }

      internal void FormatColumns() {
         if (range_ == null) {
            return;
            }
         range_[1, 1].EntireColumn.AutoFit();
         for (var ii = 4; ii <= maxMatchingColumn_; ++ii) {
            if (range_ != null) {
               range_[1, ii].EntireColumn.AutoFit();
               }
            }
         }

      internal void AddKeywords(PageNumber pageNumber, string reportText
            , string blockId, IEnumerable<string> matchingKeywords) {

         if (range_ == null) {
            return;
            }
         range_[++nextRow_, 1].Value = blockId;
         range_[nextRow_, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
         range_[nextRow_, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
         range_[nextRow_, 2].Value = pageNumber;
         range_[nextRow_, 3].Value = reportText;

         var columnIndex = 4;
         foreach (var kw in matchingKeywords) {
            maxMatchingColumn_ = Math.Max(maxMatchingColumn_, columnIndex);
            range_[nextRow_, columnIndex++].Value = kw;
            }

         // Mark this page as needing to be checked by a human
         if (!toCheck_.ContainsKey(pageNumber.PdfPageNumber)) {   
            toCheck_.Add(pageNumber.PdfPageNumber, pageNumber);
            }
         }

      internal void Finish() {
         if (range_ == null) {
            return;
            }
         var sortedPages = toCheck_.Keys.OrderBy(kk => kk)
            .Select(kk => toCheck_[kk].ToString())
            .ToList();
         range_[pageToReadRow_, 3].Value = string.Join(", ", sortedPages);
         }
      }
   }
