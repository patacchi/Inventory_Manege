using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Dapper;
using CSharp_DBHandle.CSDB_COMServer.Entity;
using CSDB_COMServer.Utility;
using SqlKata;
using SqlKata.Compilers;
using SqlKata.Execution;

namespace CSDB_COMServer
{
    public class DBUpdator<TEntity>
    where TEntity: class
    {
        private List<object[]> listarrobjValues_;
        private List<string> liststrColumns_;
        private string strTableName;

        /// <summary>
        /// コンストラクタ
        /// List<Entityクラス>を引数にとり、カラム名一覧と値のObject配列のListメンバ変数にセットする
        /// </summary>
        /// <param name="listTEntiry">エンティティクラスのList、エンティティクラスはクラス名がテーブル名になってくること</param>
        public DBUpdator(List<TEntity> listTEntiry)
        {
            if (listTEntiry is null || listTEntiry.Count() == 0)
            {
                //引数がNullもしくは長さ0のリストだった場合
                throw new ArgumentNullException();
            }
            //中身有りの場合は、クラス名がテーブル名になっているはずなのでメンバ変数テーブル名セット
            this.strTableName = typeof(TEntity).Name;
            //(Cols,Vals)のListを取得する
            DataCasting _dataCast = new DataCasting();
            var colsVals = _dataCast.getColsValuesFromEntity(listTEntiry);
            this.liststrColumns_ = colsVals.listColumuns;
            this.listarrobjValues_ = colsVals.listValues;
        }
       public async void DBUp()
        {
            ConStringBuilder conBuilder = new ConStringBuilder();
            string strConString =  conBuilder.GetACCDB_TempDBConString();
            var conFatcoty = new SqlConnectionFactory(strConString,EnumDBType.ACCDB);
            var connection =await conFatcoty.CreateConnectionAsync();
            var sqlCompiler = new SqlKata.Compilers.SqlServerCompiler();
            var db = new QueryFactory(connection,sqlCompiler);
            db.Query(strTableName)
            .AsInsert(liststrColumns_,listarrobjValues_);

            //コンストラクタで得たテーブル名とColsValsを元にクエリ構築
            var Query_ = new Query(this.strTableName)
            .AsInsert(liststrColumns_,listarrobjValues_);
            SqlResult result = sqlCompiler.Compile(Query_);
            Console.WriteLine(result.Sql);
        }
    }
}
