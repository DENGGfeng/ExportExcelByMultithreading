using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Concurrent;
using System.Data;

namespace ExportExcelByMultithreading
{
    public static class ExcelExtendtion
    {
        private static int count = 1;
        private static object lockObj = new object();

        public static void CreatExcel(DataTable dt)
        {
            IWorkbook workbook;
            workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");
            IRow header = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = header.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }
            var fileName = $@".\{DateTime.Now.Ticks}.xlsx";
            using FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            workbook.Write(fs);
        }

        public static void Insert(BlockingCollection<User> users, ISheet sheet)
        {
            while (!users.IsCompleted) { 
                var lockTaken = false;
                while (lockTaken == false) {
                    Monitor.TryEnter(lockObj, 100, ref lockTaken);
                }
                try
                {
                    User? user = null;
                    users.TryTake(out user);
                    Console.WriteLine(users.IsCompleted);
                    //Console.WriteLine(users.IsCompleted);
                    var row = sheet.CreateRow(count);
                    if (user != null)
                    {
                        row.CreateCell(0).SetCellValue(user.Id);
                        row.CreateCell(1).SetCellValue(user.Name);                     
                        Console.WriteLine(count.ToString());
                        count++;
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    Monitor.Exit(lockObj);
                }
            }
        }
    }
}
