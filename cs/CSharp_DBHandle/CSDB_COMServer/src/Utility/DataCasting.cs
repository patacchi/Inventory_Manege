using System;
using System.Reflection;

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

/*         /// <summary>
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
 */
        /// <summary>
        /// List<Entityクラス>を受け取り、リフレクションを使用してカラム一覧と値一覧のリストを返す
        /// </summary>
        /// <param name="listColumuns">List<string> カラム名一覧</param>
        /// <param name="argClass">List<object[]> 値一覧</param>
        /// <typeparam name="TEntity">プロパティに値がセットされたEntityクラス</typeparam>
        /// <returns>タプル (List<string> cols,List<object[]> values)</returns>
        public (List<string> listColumuns , List<object[]> listValues) getColsValuesFromEntity<TEntity>(List<TEntity> arglistClass)
        where TEntity:class
        {
            if (arglistClass is null)
            {
                //引数がNullだったら例外を投げる
                throw new ArgumentNullException(nameof(getColsValuesFromEntity));
            }
            //結果格納用の変数を定義
            List<string> strlistColumns_ = new List<string>();
            List<object[]> objarrlistValues_ = new List<object[]>();
            //初回判定用の変数を定義
            bool isFirst = true;
            //ListをFoeachで回す
            foreach (var elmClass in arglistClass)
            {
                //クラスのPropertyInfoを取得
                PropertyInfo[] pinfos = elmClass.GetType().GetProperties();
                //PropertyInfo[]をループ処理し、結果のListに値を設定していく
                //object[]を操作する関係上、forループの方が良い？(インデックス番号を扱いたい)
                //この時点では valsのローカルobjectの数は未確定(除外フィールドがある可能性がある)
                // object[] objarrCurrent = new object[pinfos.Count()];
                //まずはcolsにカラム名一覧を得る(初回ループ時のみ)
                if (isFirst)
                {
                    for (var varPropCounter = 0 ;varPropCounter < pinfos.Count();varPropCounter++)
                    {
                        //除外属性の有無の調査
                        var attr = pinfos[varPropCounter].GetCustomAttribute<NotIncludingValueListAttribute>();
                        if (attr is not null)
                        {
                            //除外属性がついていたら、colsには追加しないで次のループへ
                            continue;
                        }
                        //最初のみcolsにカラム名(=プロパティ名)を追加する
                        strlistColumns_.Add(pinfos[varPropCounter].Name);
                    }
                }
                //この時に、プロパティについている属性を調査し、除外対象の物はいここで落とす
                //currentObject配列の宣言 要素数はカラム名一覧のList.Count()より
                object[] objarrCurrent = new object[strlistColumns_.Count()];
                //取得したカラム名一覧をキーにして、valsを取得していく
                // カレントクラスのタイプを取得する
                Type currentType = elmClass.GetType();
                //カラム名一覧の要素分ループ
                for (var elmCounter = 0;elmCounter < strlistColumns_.Count();elmCounter++)
                {
                    objarrCurrent[elmCounter] = currentType.GetProperty(strlistColumns_.ElementAt(elmCounter))?
                    .GetValue(elmClass) ?? string.Empty;
                }
/*                 for (var varPropCounter = 0 ;varPropCounter < pinfos.Count();varPropCounter++)
                {
                    if (isFirst)
                    {
                        //除外属性の有無の調査
                        var attr = pinfos[varPropCounter].GetCustomAttribute<NotIncludingValueListAttribute>();
                        if (attr is not null)
                        {
                            //除外属性がついていたら、colsには追加しないで次のループへ
                            continue;
                        }
                        //最初のみcolsにカラム名を追加する
                        strlistColumns_.Add(pinfos[varPropCounter].Name);
                    }
                    //valsに値をセットしていく、stringがNullだった場合は String.Emptyをセットする
                    objarrCurrent[varPropCounter] =  pinfos[varPropCounter].GetValue(elmClass) ?? string.Empty;
                }
 */                //ここで1回分のcolsがセットされているはずなので、Listに追加する
                 objarrlistValues_.Add(objarrCurrent);
                //初回フラグを落とす
                isFirst = false;
            }
            return (strlistColumns_,objarrlistValues_);
        }
    }
}
