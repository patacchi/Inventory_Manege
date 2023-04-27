using System.Text.Json.Nodes;
namespace CSharp_DBHandle.CSDB_COMServer.Entitys

{
    public class JSON_Parser
    {
        JsonNode? varnode;
        /// <summary>
        ////コンストラクタ JSONファイルのパスを渡し、パースした結果をメンバ変数にセットする
        /// </summary>
        /// <param name="strJSONFilePath"></param>
        public JSON_Parser(string strJSONFilePath="")
        {
            if (strJSONFilePath == "")
            {
                //コンストラクタでパスが指定されなかった場合は定数定義されているデフォルトのパスを設定する
                strJSONFilePath = Const_Entity.DEFAULT_SETTING_JSON_PATH;
            }
           if (!File.Exists(strJSONFilePath))
            {
                //ファイルが存在しなかった場合例外を投げる
                throw new FileNotFoundException();
            }
            //まずはファイル全体を読み込む
            string strJSONRawALL = File.ReadAllText(strJSONFilePath);
            //デシリアイズを行いメンバ変数に格納(動的)
            varnode = JsonNode.Parse(strJSONRawALL);
        }

        public JsonNode? resultJsonNode{
            get
            {
                //パースしたJsonnodeをそのまま返す
                return varnode;
            }
        }
    }
}

