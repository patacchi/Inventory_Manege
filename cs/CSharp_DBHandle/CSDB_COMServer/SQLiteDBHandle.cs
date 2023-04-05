#define ACCDBMODE
#define MODE1
#define DEBUG
using System;
using System.Linq;
using System.Data.OleDb;
using FluentMigrator.Runner;
using FluentMigrator.Runner.Initialization;
using CSharp_DBHandle.CSDB_COMServer.Entity;
using Microsoft.Extensions.DependencyInjection;

namespace CSDB_COMServer

{
    public class SQLiteDBHandle
    {
        static void Main(string[] args)
        {
            CheckDB();
        }
        static void CheckDB()
        {
            using (var serviceProvider = CreateServices())
            using (var scope = serviceProvider.CreateScope())
            {
                Updatedatabase(scope.ServiceProvider);
            }
            //accdb T_INV_Label_Temp
            //接続文字列作成
            JSON_Parser jsonGlobal = new JSON_Parser();
            var jsonNodeGlobal = jsonGlobal.resultJsonNode;
            if (jsonNodeGlobal is null)
            {
                return;
            }
            if (jsonNodeGlobal["TempDBPath"] is null)
            {
                return;
            }
            //GlobalJSONからaccdbの接続文字列のひな形を取得
            System.Text.StringBuilder sbConString = new System.Text.StringBuilder();
            object[] sbParm = {Convert.ToString(jsonNodeGlobal["TempDBPath"])!};
            using (var serviceProviderAccdb = CreateServicesAccDB(sbConString.AppendFormat(Convert.ToString(jsonNodeGlobal["AccDBConString"])!, sbParm).ToString()))    
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
        private static ServiceProvider CreateServices()
        {
            return new ServiceCollection()
            //Add common FluentMigrator servives
            .AddFluentMigratorCore()
            .ConfigureRunner(rb => rb
            //Add Sqlite supoort to FluentMigrator
            .AddSQLite()
            //接続文字列作成
            .WithGlobalConnectionString("Data source=test.sqlite3")
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
