using MySql.Data.MySqlClient;
using SqlSugar;
using System.Collections.Concurrent;
using System.Data;
using System.Data.SqlClient;

namespace ExportExcelByMultithreading
{
    internal static class Data
    {
        private const string connectionStr = "user=root;password=123456;server=localhost;database=test";
        private const string sql = "select id as ID,c_name as Name from user_info";

        public static void GetData()
        {
            using MySqlConnection conn = new MySqlConnection(connectionStr);

            conn.Open();

            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            using var dataReader = cmd.ExecuteReader();
            //MySqlDataAdapter dataAdapter = new MySqlDataAdapter();
            //dataAdapter.SelectCommand = cmd;
            //DataTable dt = new DataTable();
            //dataAdapter.Fill(dt);
            //foreach(DataRow row in dt.Rows)
            //{
            //    Console.WriteLine($"{row["id"]}, {row["c_name"]}");
            //}
            while (dataReader.Read())
            {
                Console.WriteLine($"{dataReader.GetInt32("id")}-{dataReader.GetString("c_name")}");
            }
        }

        public static void GetDataBySqlSugar(BlockingCollection<User> users)
        {
            Task.Run(() =>
            {
                using SqlSugarClient client = new SqlSugarClient(new ConnectionConfig
                {
                    ConnectionString = connectionStr,
                    DbType = SqlSugar.DbType.MySql
                });
                var userReader = client.Ado.GetDataReader(sql);
                while (userReader.Read())
                {
                    ////bool res = false;
                    ////if (res == false)
                    ////{
                        var user = new User
                        {
                            Id = userReader.GetInt32(0),
                            Name = userReader.GetString(1),
                        };
                        users.Add(user);
                        //if (users.Count > 10000)
                        //{
                        //    Console.WriteLine("队列已满...");
                        //    Thread.Sleep(1000);
                        //}
                    //}
                }
                users.CompleteAdding();
                Console.WriteLine("读取数据结束。");
            });
        }
    }
}
