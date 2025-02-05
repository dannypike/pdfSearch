using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class PageNumber {
      public int PdfPageNumber { get; set; } = -1;
      public string LdsPageNumber { get; set; } = "";
      private DocumentSheet documentSheet_;

      public PageNumber(DocumentSheet documentSheet) {
         documentSheet_ = documentSheet;
         }

      public override string ToString() {
         if (0 < LdsPageNumber.Length) {
            return LdsPageNumber;
            }
         return $"PDF#{PdfPageNumber}";
         }
      }
   }