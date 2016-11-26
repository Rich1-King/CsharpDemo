using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace NopiDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Process.GetCurrentProcess().MainModule.FileName);
            Console.WriteLine(Environment.CurrentDirectory);
            ExcelUtil.writeToExcel("d:\\2.xls");
            ExcelUtil.readExcel("d:\\2.xls");
            Console.ReadKey();
        }
    }
}
