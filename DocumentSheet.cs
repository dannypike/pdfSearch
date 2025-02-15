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
      private int pagesToReadRow_;

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

         range_[++nextRow_, 3].Value = "PDF file";
         range_[nextRow_, 4].Value = Path.GetFileName(pdfFilename_);

         range_[++nextRow_, 3].Value = "Title";
         range_[nextRow_, 4].Value = "not found";
         titleRow_ = nextRow_;

         range_[++nextRow_, 3].Value = "Total Pages";
         range_[nextRow_, 4].Value = numberOfPages;
         range_[nextRow_, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

         range_[++nextRow_, 3].Value = "Pages to read";

         pagesToReadRow_ = nextRow_;

         ++nextRow_;
         ++nextRow_;
         range_[++nextRow_, 1].Value = "Matches found";
         ++nextRow_;

         // List of paragraphs where keywords were found
         firstRow_ = ++nextRow_;
         range_[nextRow_, 1].Value = "Timestamp (DP only)";
         range_[nextRow_, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
         range_[nextRow_, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
         range_[nextRow_, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

         range_[nextRow_, 2].Value = "ID (DP only)";
         range_[nextRow_, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
         range_[nextRow_, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
         range_[nextRow_, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

         range_[nextRow_, 3].Value = "Page";
         range_[nextRow_, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

         range_[nextRow_, 4].Value = "Context";
         range_[nextRow_, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

         range_[nextRow_, 5].Value = "Keywords";
         range_[nextRow_, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
         }

      internal void SetTitle(string title) {
         titleText_ = title;
         if (range_ != null) {
            range_[titleRow_, 4].Value = titleText_;
            }
         }

      internal void FormatColumns() {
         if (range_ == null) {
            return;
            }
         for (var ii = 1; ii <= maxMatchingColumn_; ++ii) {
            range_[1, ii].EntireColumn.AutoFit();
            }
         }

      internal void AddKeywords(PageNumber pageNumber, string reportText
            , string blockId, IEnumerable<string> matchingKeywords) {

         if (range_ == null) {
            return;
            }

         range_[++nextRow_, 1].Value = $"{Program.Timestamp:dd MMM yyyy HH:mm:ss}";
         range_[nextRow_, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
         range_[nextRow_, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);

         range_[nextRow_, 2].Value = blockId;
         range_[nextRow_, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
         range_[nextRow_, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);

         range_[nextRow_, 3].Value = pageNumber;
         range_[nextRow_, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

         range_[nextRow_, 4].Value = reportText;

         // Mark this page as needing to be checked by a human
         if (!toCheck_.ContainsKey(pageNumber.PdfPageNumber)) {   
            toCheck_.Add(pageNumber.PdfPageNumber, pageNumber);
            }

         // And list the keywords that matched
         var column = 4;
         foreach (var mkw in matchingKeywords) {
            range_[nextRow_, ++column].Value = mkw ?? "";
            range_[nextRow_, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            // Remember the maximum number of column in this sheet for auto-fit later
            maxMatchingColumn_ = Math.Max(maxMatchingColumn_, column);
            }
         }

      internal void Finish() {
         if (range_ == null) {
            return;
            }

         range_[pagesToReadRow_, 4].Value = toCheck_.Count;
         
         // Combine adjacent pages into ranges
         var checkCount = toCheck_.Count;
         if (checkCount == 0) {
            // degenerate case
            return;
            }
         if (checkCount == 2) {
            // simple cases
            var twoPages = toCheck_.Keys.OrderBy(kk => kk).Select(kk => toCheck_[kk].ToString());
            range_[pagesToReadRow_, 4].Value = string.Join(", ", twoPages);
            return;
            }

         // Convert individual pages into tuples with min/max values that are equal
         var minMax = toCheck_.Keys.OrderBy(kk => kk).Select(kk => (MinPage: toCheck_[kk], MaxPage: toCheck_[kk])).ToList();
         var pageCount = minMax.Count;
         range_[pagesToReadRow_ + 1, 4].Value = $"(a total of {pageCount} {Program.pluralled("page", pageCount)})";

         // Combine adjacent entries, if they are adjacent in the PDF page numbering
         var index = 0;
         while (index < minMax.Count - 1) {
            var thisPage = minMax[index].MaxPage;
            var nextPage = minMax[index + 1].MinPage;
            if (thisPage.PdfPageNumber + 1 == nextPage.PdfPageNumber) {
               minMax[index] = (MinPage: minMax[index].MinPage, MaxPage: minMax[index+1].MaxPage);
               minMax.RemoveAt(index + 1);
               }
            else {
               ++index;
               thisPage = nextPage;
               }
            }

         // Then display those ranges
         var csvPages = minMax
            .Select(pp => pp.MinPage == pp.MaxPage ? pp.MinPage.ToString() : $"{pp.MinPage}-{pp.MaxPage}");
         range_[pagesToReadRow_, 4].Value = string.Join(", ", csvPages);

         // Auto0fit the keyword columns
         for (var column = 5; column <= maxMatchingColumn_; ++column) {
            range_[1, column].EntireColumn.AutoFit();
            }
         }
      }
   }
