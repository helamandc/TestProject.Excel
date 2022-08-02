using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject.Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {

            var EditNow = new EditSpreadSheet();

            EditNow.Edit();



                //IWorkbook workbook = new XSSFWorkbook();
                //ISheet worksheet = workbook.CreateSheet();
                //IRow row = worksheet.CreateRow(0);
                //ICell cell = row.CreateCell(2);
                //cell.SetCellValue("HelloWorld");
                //using (var fileData = new FileStream(@"C:\temp\test.xlsx", FileMode.Create))
                //{
                //    workbook.Write(fileData);
                //}
            
            

            
                
            
        }

           
    }
}
