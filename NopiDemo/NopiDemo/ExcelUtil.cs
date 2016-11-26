using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NopiDemo
{
    public class ExcelUtil
    {
        public static void writeToExcel(string fileName)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet1");
            sheet.SetColumnWidth(0, 15*256);
            sheet.SetColumnWidth(1, 20*256);
            sheet.SetColumnWidth(2, 15*256);
            for (int i=0; i<10; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.Height = 30;
                if(i == 0)
                {
                    ICell cellHeader1 = row.CreateCell(0);
                    ICell cellHeader2 = row.CreateCell(1);
                    ICell cellHeader3 = row.CreateCell(2);
                    cellHeader1.SetCellValue("编号");
                    cellHeader2.SetCellValue("名字");
                    cellHeader3.SetCellValue("随机数");
                    continue;
                }
                ICell cell1 = row.CreateCell(0);
                cell1.SetCellValue(i);
                ICell cell2 = row.CreateCell(1);
                cell2.SetCellValue("张三");
                ICell cell3 = row.CreateCell(2);
                cell3.SetCellValue(new Random().Next(100, 900));
            }
            FileStream fs = null;
            try{
                fs = File.OpenWrite(fileName);
                wb.Write(fs); //向打开的xls文件写入内存中的wookbook 
            }catch(Exception e){
                Console.WriteLine("写入失败"+e);
            }finally {
               fs.Close(); 
               wb.Close();
            }
        }

        public static void readExcel(string fileName)
        {
            FileStream fs = null;
            HSSFWorkbook wb = null;
            try {
                fs = File.OpenRead(fileName);
                wb = new HSSFWorkbook(fs);
                for(int i=0; i<wb.NumberOfSheets; i++)
                {
                    ISheet sheet = wb.GetSheetAt(i);
                    for(int j=0; j<sheet.LastRowNum; j++)
                    {
                        IRow row = sheet.GetRow(j);
                        for(int k=0; k<row.LastCellNum; k++)
                        {
                            ICell cell = row.GetCell(k);
                            if(cell != null)
                            {
                                Console.Write(cell.ToString()+"        ");
                            }
                        }
                        Console.WriteLine();
                    }
                }
            }catch(Exception e) {
                Console.WriteLine("读取excel失败"+e);
            }finally {
               fs.Close();
               wb.Close();
            }
        }
    }
}
