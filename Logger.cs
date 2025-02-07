using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PdfSearch {
   internal class Logger {
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
          static extern void OutputDebugStringW(string lpOutputString);

      public static void WriteLine(string message, bool copyConsole = true) {
         //OutputDebugStringW($"{message}\r\n");
         Debug.WriteLine(message);
         if (copyConsole) {
            Console.WriteLine($"\r\u001b[K\r{message}");
            }
         }
      }
   }
