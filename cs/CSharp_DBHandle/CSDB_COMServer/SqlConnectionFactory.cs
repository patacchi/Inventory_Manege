using System.Data;

namespace CSDB_COMServer
{
    /// <summary>
    /// Connection作成に使うインターフェース
    /// </summary>
    public interface IDatabaseConnectionFactory
    {
        Task<IDbConnection> CreateConnectionAsync();
    }
    public enum EnumDBType
    {
        SQLite = 1,
        ACCDB = 2,
        SQLLocalDB = 3
    }
    /// <summary>
    /// 
    /// </summary>
    public class SqlConnectionFactory:IDatabaseConnectionFactory
    {
        private readonly string _connectionString;
        private readonly EnumDBType _dbTypeEnum;

        /// <summary>
        /// コンストラクション、メンバ変数に接続文字列をセットする
        /// </summary>
        /// <param name="connectionString">接続文字列を渡す</param>
        /// <returns>接続文字列がNullの時はNullExceptionを投げる</returns>
        public SqlConnectionFactory(string connectionString,EnumDBType dbTypeEnum)
        { 
            _connectionString = connectionString ??
            throw new ArgumentNullException(nameof(connectionString));
            _dbTypeEnum = dbTypeEnum;
            return;
        }

        /// <summary>
        /// Connectionオブジェクトを返す。EnumDBTypeの指定が必須
        /// </summary>
        /// <param name="dbTypeEnum">DBTypeの列挙型</param>
        /// <returns></returns>
        public async Task<IDbConnection> CreateConnectionAsync()
        {
            switch (_dbTypeEnum)
            {
                /// <summary>
                /// SQLiteの場合
                /// </summary>
                case EnumDBType.SQLite:
                {
                    var sqlConnection = new System.Data.SQLite.SQLiteConnection(_connectionString);
                    await sqlConnection.OpenAsync();
                    return sqlConnection;
                }
                /// <summary>
                /// ACCDBの場合
                /// </summary>
                case EnumDBType.ACCDB:
                {
                    var sqlConnection = new System.Data.OleDb.OleDbConnection(_connectionString);
                    await sqlConnection.OpenAsync();
                    return sqlConnection;
                }
                /// <summary>
                /// SQL LocalDBの場合(工事中)
                /// </summary>
                /// <returns></returns>
                case EnumDBType.SQLLocalDB:
                {
                    throw new NotImplementedException();
                }
                default:
                {
                    throw new ArgumentNullException();
                }
            }
        }
    }
}
