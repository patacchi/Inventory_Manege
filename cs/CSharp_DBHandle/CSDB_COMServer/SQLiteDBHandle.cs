#define ACCDBMODE
#define MODE1
#define DEBUG
using System;
using System.Linq;
using System.Data.OleDb;
using FluentMigrator.Runner;
using FluentMigrator.Runner.Initialization;
using CSDB_COMServer.Entitys;
using Microsoft.Extensions.DependencyInjection;

namespace CSDB_COMServer

{
    public class SQLiteDBHandle
    {
        static void Main(string[] args)
        {
            // var SQLiteH = new SQLiteDBHandle();
            CheckDB();
        }
        public static void CheckDB()
        {
            ConStringBuilder conBuilder = new ConStringBuilder();         
            using (var serviceProvider = CreateServices(conBuilder.GetSqlite_TempDBConString()))
            using (var scope = serviceProvider.CreateScope())
            {
                Updatedatabase(scope.ServiceProvider);
            }
            //接続文字列取得
            // ConStringBuilder conBuilder = new ConStringBuilder();         
            using (var serviceProviderAccdb = CreateServicesAccDB(conBuilder.GetACCDB_TempDBConString()))    
            using ( var scopeAccdb = serviceProviderAccdb.CreateScope())
            {
                Updatedatabase(scopeAccdb.ServiceProvider);
            }
            return;
        }
        /// <summary>
        /// Dependency Injection 初期設定
        /// </summary>
        /// <returns></returns>
        private static ServiceProvider CreateServices(string strConnection)
        {
            return new ServiceCollection()
            //Add common FluentMigrator servives
            .AddFluentMigratorCore()
            .ConfigureRunner(rb => rb
            //Add Sqlite supoort to FluentMigrator
            .AddSQLite()
            //接続文字列作成
            // .WithGlobalConnectionString("Data source=test.sqlite3")
            .WithGlobalConnectionString(strConnection)
            //マイグレーションに使用するアセンブリを指定する
            .ScanIn(typeof(T_INV_Label_Temp).Assembly).For.Migrations())
        // コンソールログ有効化
        .AddLogging(lb => lb.AddFluentMigratorConsole())
        //Build service provider
        .BuildServiceProvider(false);
        }
        private static ServiceProvider CreateServicesAccDB(string strConnection)
        {
            return new ServiceCollection()
            //Add common FluentMigrator servives
            .AddFluentMigratorCore()
            .ConfigureRunner(rb => rb
            //Add Jet supoort to FluentMigrator
            .AddJet()
            //接続文字列作成
            .WithGlobalConnectionString(strConnection)
            // (@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\cs\CSharp_DBHandle\CSDB_COMServer\test_Local.accdb")
            //マイグレーションに使用するアセンブリを指定する
            .ScanIn(typeof(T_INV_Label_Temp).Assembly).For.Migrations())
        // コンソールログ有効化
        .AddLogging(lb => lb.AddFluentMigratorConsole())
        //Build service provider
        .BuildServiceProvider(false);
        }

        private static void Updatedatabase(IServiceProvider serviceProvider)
        {
            //Instantiate the runner
            var runner = serviceProvider.GetRequiredService<IMigrationRunner>();

            //Execute the migrations
            try
            {
                runner.MigrateUp();
            }
            catch (OleDbException olex)
            {
                System.Text.StringBuilder sbError;
                sbError = new System.Text.StringBuilder();
                for (int iErrCount = 0 ; iErrCount < olex.Errors.Count;iErrCount++)
                {
                    sbError.AppendLine(olex.Errors[iErrCount].Message);
                }
                Console.WriteLine(sbError.ToString());
            }
            catch (System.Exception exceptione)
            {
                Console.WriteLine(exceptione.Message);
            }
            
        }
        private static void UpdatedatabaseAccdb(IServiceProvider serviceProvider)
        {
            //Instantiate the runner
            var runner = serviceProvider.GetRequiredService<IMigrationRunner>();

            //Execute the migrations
            runner.MigrateUp();
        }

    }
}
