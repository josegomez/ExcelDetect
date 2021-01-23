using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDetect
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Detected:"+Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Excel.Application").ToString());
                Console.Read();
               }
            catch(Exception e )
            {
                Console.WriteLine("Not Found: " + e.ToString());
                Console.Read();
            }
        }
    }
}
