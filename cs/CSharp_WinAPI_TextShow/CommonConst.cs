using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
///<summary>
///共通参照定数を集めたNameSpace 
///</summary>
namespace CommonConst
{
    /// <summary>
    /// DBに関する定数・定義を集めたNameSpace
    /// </summary>
    namespace DBConst
    {
        /// <summary>
        /// DB共通の定数を参照するクラス
        /// </summary>
        public static class DBCommon
        {
            private readonly static string strDefaultDBPath = @"C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB";
            /// <summary>
            /// デフォルトDBフルパスの定数
            /// </summary>
            public static string StrDefaultDBPath => strDefaultDBPath;
            //private readonly static string strDBFileName = @"INV_Manege.accdb";
            /// <summary>
            /// デフォルトDBファイル名の定数
            /// </summary>
            public static string StrDBFileName { get; }  = @"INV_Manege.accdb";
            private readonly static string strTempDBFileName = @"DB_Temp_Local.accdb";
            /// <summary>一時DBファイル名の定数</summary>
            public static string StrTempDBFileName => strTempDBFileName;
        }
    }
}
namespace CSharp_WinAPI_TextShow
{
    /// <summary>
    /// テスト実装クラス
    /// </summary>
    class CommonCont
    {
    }
}
