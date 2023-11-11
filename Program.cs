using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SqlSugar;
using System.Collections.Concurrent;
using System.Diagnostics;

internal class Program
{
    static int count = 1;
    static object lockObj = new object();
    private const string connectionStr = "user=root;password=123456;server=localhost;database=test";
    private const string sql = "select id as ID,c_name as Name from user_info";
    static BlockingCollection<User> users = new BlockingCollection<User>();
    private static async Task Main(string[] args)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        stopwatch.Start();
        Console.WriteLine("开始导出...");

        IWorkbook workbook;
        workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Sheet1");
        IRow header = sheet.CreateRow(0);
        ICell headerCell0 = header.CreateCell(0);
        headerCell0.SetCellValue("ID");
        ICell headerCell1 = header.CreateCell(1);
        headerCell1.SetCellValue("Name");

        GetDataBySqlSugat();

        List<Task> tasks = new()
        {
            Task.Run(() => Insert(users, sheet)),
            Task.Run(() => Insert(users, sheet)),
            Task.Run(() => Insert(users, sheet)),
            Task.Run(() => Insert(users, sheet)),
            Task.Run(() => Insert(users, sheet)),
        };

        await Task.WhenAll(tasks);

        var fileName = $@".\{DateTime.Now.Ticks}.xlsx";
        using FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(fs);
        Console.WriteLine("导出完成。");
        stopwatch.Stop();
        Console.WriteLine($"耗时：{stopwatch.ElapsedMilliseconds / 1000}s");
    }

    private static void Insert(BlockingCollection<User> users, ISheet sheet)
    {
        while (!users.IsCompleted)
        {
            users.TryTake(out User? user);
            lock (lockObj) {
                var row = sheet.CreateRow(count);
                if (user != null)
                {
                    row.CreateCell(0).SetCellValue(user.Id);
                    row.CreateCell(1).SetCellValue(user.Name);
                    Console.WriteLine(count.ToString());
                    count++;
                }
            }
        }
    }

    public static void GetDataBySqlSugat()
    {
        Task.Run(() =>
        {
            using SqlSugarClient client = new SqlSugarClient(new ConnectionConfig
            {
                ConnectionString = connectionStr,
                DbType = DbType.MySql
            });
            var list = client.SqlQueryable<User>(sql).ToList();
            foreach (var item in list)
            {
                users.Add(item);
            }
            //var userReader = client.Ado.GetDataReader(sql);
            //while (userReader.Read())
            //{
            //    //bool res = false;
            //    //if (res == false)
            //    //{
            //        var user = new User
            //        {
            //            Id = userReader.GetInt32(0),
            //            Name = userReader.GetString(1),
            //        };
            //        users.Add(user);
            //        //if (users.Count > 20000)
            //        //{
            //        //    Console.WriteLine("队列已满...");
            //        //    Thread.Sleep(1000);
            //        //}
            //    //}
            //}
            users.CompleteAdding();
            Console.WriteLine("读取数据结束。");
        });
    }
}