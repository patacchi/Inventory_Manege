using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Dapper;
using CSharp_DBHandle.CSDB_COMServer.Entity;
using SqlKata;
using SqlKata.Compilers;
using SqlKata.Execution;

namespace CSDB_COMServer
{

    public class DBUpdator<TEntity>
    where TEntity: class
    {
        
        private dynamic _arrResult{get;}

        public DBUpdator(List<TEntity> listTEntiry)
        {
                        
        }

        public DBUpdator(dynamic arrResult)
        {
            if (arrResult.Length == 0)
            {
                //引数がnullだった
                Console.WriteLine(nameof(arrResult).ToString() + " arg is null.");
                throw new ArgumentNullException();
            }
            this._arrResult = arrResult;
            return;
        }
        public async void DBUp()
        {
            ConStringBuilder conBuilder = new ConStringBuilder();
            string strConString =  conBuilder.GetACCDB_TempDBConString();
            var conFatcoty = new SqlConnectionFactory(strConString,EnumDBType.ACCDB);
            var connection =await conFatcoty.CreateConnectionAsync();
            var sqlCompiler = new SqlKata.Compilers.SqlServerCompiler();
            // var db = new QueryFactory(connection,sqlCompiler);
            var Query_ = new Query(_arrResult[0].GetType().Name)
            .AsInsert(_arrResult[0]);
            SqlResult result = sqlCompiler.Compile(Query_);
            Console.WriteLine(result.Sql);
/*            var db = new QueryFactory(connection,sqlCompiler);
             db.Query(_arrResult[0].GetType().Name.ToString())
            .Insert(new {F_INV_TehaiCode = "3A8A1314P010"});
            db.Logger = compiled =>
               Console.WriteLine(compiled.Sql);
 */        }
    }
}
