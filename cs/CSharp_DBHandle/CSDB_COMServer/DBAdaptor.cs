using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Dapper;
using CSDB_COMServer.Entitys;
using CSDB_COMServer.Utility;
using SqlKata;
using SqlKata.Compilers;
using SqlKata.Execution;

namespace CSDB_COMServer
{
    public class EntityUpdator<TEntity>
    where TEntity: class
    {
        private List<TEntity> listEntity_;
        private List<object[]> listarrobjValues_;
        private List<string> liststrColumns_;
        private string strTableName;

        /// <summary>
        /// コンストラクタ
        /// List<Entityクラス>を引数にとり、カラム名一覧と値のObject配列のListメンバ変数にセットする
        /// </summary>
        /// <param name="listTEntiry">エンティティクラスのList、エンティティクラスはクラス名がテーブル名になってくること</param>
        public EntityUpdator(List<TEntity> listTEntiry)
        {
            if (listTEntiry is null || listTEntiry.Count() == 0)
            {
                //引数がNullもしくは長さ0のリストだった場合
                throw new ArgumentNullException();
            }
            //エンティティクラスのListをメンバ変数へセット
            listEntity_ = listTEntiry;
            //中身有りの場合は、クラス名がテーブル名になっているはずなのでメンバ変数テーブル名セット
            this.strTableName = typeof(TEntity).Name;
            //ここでテーブル・フィールド存在チェック(無ければ作成)する？
            SQLiteDBHandle.CheckDB();
            //(Cols,Vals)のListを取得する
            DataCasting _dataCast = new DataCasting();
            var colsVals = _dataCast.getColsValuesFromEntity(listTEntiry);
            this.liststrColumns_ = colsVals.listColumuns;
            this.listarrobjValues_ = colsVals.listValues;
        }
       public async void DBUp(EnumDBType dbTypeEnum_)
        {
            ConStringBuilder conBuilder = new ConStringBuilder();
            // string strConString =  conBuilder.GetACCDB_TempDBConString();
            string strConString =  conBuilder.GetSqlite_TempDBConString();
            // var conFatcoty = new SqlConnectionFactory(strConString,EnumDBType.ACCDB);
            var conFatcoty = new SqlConnectionFactory(strConString,EnumDBType.SQLite);
            var connection =await conFatcoty.CreateConnectionAsync();
            var sqlCompiler = new SqlKata.Compilers.SqliteCompiler();
            //コンストラクタで得たテーブル名とColsValsを元にクエリ構築
            var Query_ = new Query(this.strTableName)
            .AsInsert(liststrColumns_,listarrobjValues_);
            SqlResult result = sqlCompiler.Compile(Query_);
            /*
            #if DEBUG
            {
                Console.WriteLine(result.Sql);
            }
            #endif
            */
            // 得られたSQLを利用して、INSERT実行
            //パラメータリストを作成する
            var parms = new DynamicParameters();
            for (int intCounter = 0;intCounter <= result.Bindings.Count() -1 ;intCounter++)
            {
                //パラメータ名は p{連番} となっている
                parms.Add($"p{intCounter.ToString()}",result.Bindings[intCounter]);
            }
            var rows = connection.Execute(result.Sql,parms);

        }
    }
}
