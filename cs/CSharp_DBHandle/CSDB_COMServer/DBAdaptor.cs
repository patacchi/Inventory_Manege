using System;
using CSharp_DBHandle.CSDB_COMServer.Entity;
using DbExtensions;

namespace CSDB_COMServer
{
    /// <summary>
    /// Entity モデルクラスを受け取り、DBにアップデートするクラス
    /// </summary>
    public class DBUpdator
    {
        public DBUpdator(object objEntity)
        {
            if (objEntity is null)
            {
                //引数がnullだった
                Console.WriteLine(nameof(DBUpdator).ToString() + " arg is null.");
                throw new ArgumentNullException();
            }
            SqlBuilder sqlQuery = new SqlBuilder();
            sqlQuery = SQL
            .SELECT(nameof(T_INV_Label_Temp.F_InputDate))
            ._(nameof(T_INV_Label_Temp.F_INV_Current_Amount))
            .FROM(nameof(T_INV_Label_Temp));
        }
    }
}
