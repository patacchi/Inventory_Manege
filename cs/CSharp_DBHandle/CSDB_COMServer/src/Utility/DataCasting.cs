using System;

namespace CSDB_COMServer.Utility
{
    public class DataCasting
    {
        private object[,] _arrObj2Dim;

        /// <summary>
        /// コンストラクタ。今後何かやる？
        /// </summary>
        public DataCasting()
        {
            _arrObj2Dim = new object[0,0];
            return;
        }

        /// <summary>
        /// List<Entirty>から Object[] に変換するメソッド
        /// </summary>
        /// <param name="arrayArg"></param>
        /// <returns>Object[]</returns>
        public object[] castArrayToObject(Array arrayArg)
        {
            if (arrayArg is null || arrayArg.Length == 0)
            {
                //長さ0の配列が来たら引数無しとして例外を投げる
                throw new ArgumentNullException(nameof(castArrayToObject));
            }
            if (arrayArg.Rank != 2)
            {
                //2次元以外の配列が来たら今のところは未実装にし、例外を投げる
                throw new NotImplementedException(nameof(this.castArrayToObject));
            }
            //引数の要素数分のObject配列を宣言(ローカル)
            object[,] objLocal = new object[arrayArg.Length-1,1];
            _arrObj2Dim = objLocal;
            return new object[0];
        }
    }
}
