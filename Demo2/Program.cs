using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Demo2
{
    class Program
    {
        static Excel.Application _xlApp = new Excel.Application();
        static string _pDebug = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        static List<User> _users = new List<User>();

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            ReadData(@"D:\Book\Customer.xlsx");
            WriteData(@"H:\TestFolder");
        }

        static void ReadData(string psourceFileName)
        {
            if (!File.Exists(psourceFileName))
            {
                Console.WriteLine("***File của bạn không tồn tại.");
                return;
            }

            Excel.Workbook xlWb = _xlApp.Workbooks.Open(psourceFileName);
            Excel.Worksheet excelSheet = xlWb.ActiveSheet;

            //count row
            var lastRow = excelSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            var lastColumn = excelSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            var rowStart = 3;
            for (int i = rowStart; i <= lastRow; i++)
            {
                var u = new User
                {
                    UserId = excelSheet.Cells[i, 5].Value.ToString(),
                    Name = excelSheet.Cells[i, 2].Value.ToString(),
                    Phone = excelSheet.Cells[i, 4].Value.ToString(),
                };
                _users.Add(u);
            }

            xlWb.Close(0);
        }

        static void WriteData(string pdest)
        {
            var sourceFileName = _pDebug + @"/template_v2.xlsx";
            var destFileName = pdest + $"//BaoCaoHangNgay_{DateTimeToString()}.xlsx";

            //check datafile
            if (_users == null || _users.Count == 0)
            {
                Console.WriteLine("***File của bạn không chứa dữ liệu.");
                return;
            }

            //create directory
            if (!Directory.Exists(pdest))
                Directory.CreateDirectory(pdest);

            //copy file template vao duong dan
            File.Copy(sourceFileName, destFileName);

            //edit file
            Excel.Workbook xlWb = _xlApp.Workbooks.Open(destFileName);
            Excel.Worksheet excelSheet = xlWb.ActiveSheet;

            //read/write the datetime cell
            string fieldTime = excelSheet.Cells[2, 1].Value.ToString();
            excelSheet.Cells[2, 1] = fieldTime + DateTime.Now.ToString(" dd/MM/yyyy");

            //read/write name author
            string fieldAuthor = excelSheet.Cells[4, 1].Value.ToString();
            excelSheet.Cells[4, 1] = fieldAuthor + " Vũ Thị Duyên";

            #region write content
            var rowStart = 7;
            var index = 1;
            //uppercase
            TextInfo textInfo = Thread.CurrentThread.CurrentCulture.TextInfo;

            foreach (var user in _users)
            {
                excelSheet.Cells[rowStart, 1] = index++;
                excelSheet.Cells[rowStart, 2] = user.UserId;
                excelSheet.Cells[rowStart, 3] = textInfo.ToTitleCase(user.Name.ToLower());
                excelSheet.Cells[rowStart, 4] = "'0" + user.Phone;
                rowStart++;
            }

            #endregion

            //center
            for (int i = 1; i <= 4; i++)
            {
                string c = ((char)(i + 64)).ToString();
                excelSheet.get_Range(c + 7, c + rowStart).Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            }

            //save and close file
            xlWb.Save();
            xlWb.Close();
            Console.WriteLine("Thành công!!!");
        }

        static string DateTimeToString()
        {
            var now = DateTime.Now;
            return now.ToString("ddMMyyyy_hhmmss");
        }
    }
}
