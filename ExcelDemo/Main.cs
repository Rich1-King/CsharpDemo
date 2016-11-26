using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelDemo
{
    class MainController
    {
        static void Main(string[] args)
        {
            ExcelUtil.writeExcel("d:\\test.xls");
            ExcelUtil.readExcel("d:\\test.xls");
        }
    }
}
