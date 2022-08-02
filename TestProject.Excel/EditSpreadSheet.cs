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
    internal class EditSpreadSheet
    {
        public void Edit()
        {
            string pathSource = @"C:\temp\test.xlsx";

            IWorkbook templateWorkbook;
            using (FileStream fs = new FileStream(pathSource, FileMode.Open, FileAccess.Read))
            {
                templateWorkbook = new XSSFWorkbook(fs);
            }

            string sheetName = "ImportTemplate";
            ISheet sheet = templateWorkbook.GetSheet(sheetName) ?? templateWorkbook.CreateSheet(sheetName);
            IRow dataRow = sheet.GetRow(4) ?? sheet.CreateRow(4);
            ICell cell = dataRow.GetCell(1) ?? dataRow.CreateCell(1);
            cell.SetCellValue("foo");

            using (FileStream fs = new FileStream(pathSource, FileMode.Create, FileAccess.Write))
            {
                templateWorkbook.Write(fs);

            }
        }
    }
}
