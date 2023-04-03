#define ACCDBMODE
#define MODE1
#define DEBUG
using System;
using System.Linq;
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
            using (var serviceProvider = CreateServices())
            using (var scope = serviceProvider.CreateScope())
            {
                Updatedatabase(scope.ServiceProvider);
            }
            using (var serviceProviderAccdb = CreateServicesAccDB())
            using ( var scopeAccdb = serviceProviderAccdb.CreateScope())
            {
                Updatedatabase(scopeAccdb.ServiceProvider);
            }
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
        private static ServiceProvider CreateServicesAccDB()
        {
            return new ServiceCollection()
            //Add common FluentMigrator servives
            .AddFluentMigratorCore()
            .ConfigureRunner(rb => rb
            //Add Sqlite supoort to FluentMigrator
            .AddJet()
            //接続文字列作成
            .WithGlobalConnectionString("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=test.accdb")
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
            runner.MigrateUp();
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
