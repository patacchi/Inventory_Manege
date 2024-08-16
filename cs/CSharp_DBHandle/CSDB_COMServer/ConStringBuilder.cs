using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CSDB_COMServer.Entitys;

namespace CSDB_COMServer
{
    public class ConStringBuilder
    {
        JSON_Parser jsonGlobal;


        /// <summary>
        /// コンストラクター メンバ変数に JSONファイルをセットする
        /// </summary>
        /// <param name="strJsonPath">JSONファイルの場所を指定(オプション)</param>
        public ConStringBuilder(string strJsonPath = "")
        {
            if (strJsonPath == "")
            {
                jsonGlobal = new JSON_Parser();
                return;
            }
            else
            {
                if (!File.Exists(strJsonPath))
                {
                    //ファイル指定されているが、見つからなかった場合
                    //デフォルトのパスを使用する
                    Console.WriteLine("File Not Found. Use Default JSON Path");
                    jsonGlobal = new JSON_Parser();
                    return;
                }
                //JSONファイルが指定されている場合はそれを使用する
                jsonGlobal = new JSON_Parser(strJsonPath);
                return;
            }
        }
        /// <summary>
        /// Oledb を使用する際の accdbのTemopテーブルファイルへの接続文字列を返す
        /// </summary>
        /// <returns></returns>
        public string GetACCDB_TempDBConString()
        {
            var jsonNodeGlobal = jsonGlobal.resultJsonNode;
            //null チェック
            if (jsonNodeGlobal is null)
            {
                return (string.Empty);
            }
            if ((jsonNodeGlobal["TempDBPath"] is null) || ((jsonNodeGlobal["AccDBConString"]) is null))
            {
                return (string.Empty);
            }
            //GlobalJson よりaccdb接続文字列のひな型を取得
            System.Text.StringBuilder sbAccdb = new System.Text.StringBuilder();
            //パラメータ置換用の配列を準備(accdbファイルパス)
            object[] arrParm = {Convert.ToString(jsonNodeGlobal["TempDBPath"])!};
            //ひな形のパラメータ置換して、結果として返す
            return sbAccdb.AppendFormat(Convert.ToString(jsonNodeGlobal["AccDBConString"])!,arrParm).ToString();
        }

        public string GetSqlite_TempDBConString()
        {
            var jsonNodeGlobal = jsonGlobal.resultJsonNode;
            //null チェック
            if (jsonNodeGlobal is null)
            {
                return (string.Empty);
            }
            if ((jsonNodeGlobal["SqliteTempDBPath"] is null) || ((jsonNodeGlobal["SqliteConString"]) is null))
            {
                return (string.Empty);
            }
            //置換操作用の StringBuilderを用意する
            System.Text.StringBuilder sbSqlite = new System.Text.StringBuilder();
            //パラメータ置換用の配列を準備(Sqliteファイルパス)
            object[] arrParm = {Convert.ToString(jsonNodeGlobal["SqliteTempDBPath"])!};
            //ひな形のパラメータ置換を実行して、結果として返す
            return sbSqlite.AppendFormat(Convert.ToString(jsonNodeGlobal["SqliteConString"])!,arrParm).ToString();
            // return new NotImplementedException().ToString();
        }
    }
}