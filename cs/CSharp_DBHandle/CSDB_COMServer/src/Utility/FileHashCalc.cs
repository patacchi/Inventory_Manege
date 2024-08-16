using System;
using System.IO;
using System.Security.Cryptography;
namespace CSDB_COMServer.Utility
{
    static public class FileHashCalc
    {
        /// <summary>
        /// ファイルパスよりSHA512ハッシュを求める
        /// </summary>
        /// <param name="filePath_">ファイルパス</param>
        /// <returns></returns>
        public static string CreateSHA512String(string filePath_)
        {
            //nullだった
            if (filePath_ is null)
            {
                return new ArgumentNullException().ToString();
            }
            //ファイルが存在しなかった
            if (!File.Exists(filePath_))
            {
                System.Windows.Forms.MessageBox.Show(filePath_ +" は存在しませんでした");
                return new FileNotFoundException().ToString();
            }
            using (FileStream fs = new FileStream(filePath_,FileMode.Open,FileAccess.Read,FileShare.Read))
            {
                return CreateSHA512String(fs);
            }
        }
        /// <summary>
        /// FileStreamよりSHA512ファイルハッシュを求める
        /////// </summary>
        /// <param name="fs_">FileStream</param>
        /// <returns>ハイフンなしの16進文字列</returns>
        public static string CreateSHA512String(FileStream fs_)
        {
            SHA512 provideSHA512 = SHA512.Create();
            //バイト配列として結果を受け取る
            byte[] bsSHA = provideSHA512.ComputeHash(fs_);
            //バイト配列を文字列に変換し、戻り値とする
            return BitConverter.ToString(bsSHA).Replace("-","");
        }
    }
}
