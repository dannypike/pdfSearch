using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class MatchedKeyword {
      public int Count { get; private set; }

      public string Keyword { get; private set; } = "";
      public HashSet<string> Matches { get; private set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

      public int AddCount(string matchedWord, int increment = 1) {
         Matches.Add(matchedWord);
         Count += increment;
         return Count;
         }
      }
   }
