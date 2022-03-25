using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharp_WinAPI_TextShow
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length <= 0)
            {
                //引数が無かったら終了する
                MessageBox.Show(@"コマンドライン引数が必要です。引数としてラベルファイルのパスを指定して下さい" +
                                @"カレントディレクトリ:" + System.IO.Directory.GetCurrentDirectory());
                return;
            }
            if  (!System.IO.File.Exists(args[0]))
            {
                MessageBox.Show("引数で与えられたファイルが存在しませんでした");
                return;
            }
            else
            {
                string strExtention = System.IO.Path.GetExtension(args[0]).ToLower();
                switch (strExtention)
                {
                    case ".txt":
                        //txtファイルだったら
                        //メモ帳のプロセス情報を取得、引数はファイルパス
                        //System.Diagnostics.ProcessStartInfo psiNotepad = new System.Diagnostics.ProcessStartInfo(@"notepad.exe", " \"" + args[0] + "\"");
                        System.Diagnostics.ProcessStartInfo psiNotepad = new System.Diagnostics.ProcessStartInfo(@"notepad.exe", " \"" + System.IO.Path.GetFullPath(args[0]) + "\"");
                        System.Diagnostics.Process.Start(psiNotepad);
                        break;
                    case ".zip":
                        //zipファイルだったら
                        //Web上からDLするのでファイルパス表示テスト
                        MessageBox.Show(System.IO.Path.GetFullPath(args[0]));
                        //System.Diagnostics.ProcessStartInfo psiExplzh = new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\Explzh\EXPLZH.EXE", "" + System.IO.Path.GetFullPath(args[0]) + "");
                        System.Diagnostics.ProcessStartInfo psiExplzh = new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\Explzh\EXPLZH.EXE", "\"" + args[0] + "\"");
                        System.Diagnostics.Process.Start(psiExplzh);
                        break;
                    case ".lbl":
                        //lblファイルだったら
                        //本来のラベル出力プログラムに渡して、並列してDBに格納する処理を実装する
                        //\HKEY_CLASSES_ROOT\.LBL
                        //(既定) REZ_SZ "lbl_auto_file"
                        //以下のみ変更
                        //元々の設定の場所 \HKEY_CLASSES_ROOT\LBL_auto_file\shell\open\command
                        //値："C:\Program Files\用品管理ラベル出力\label002.exe" %1
                        //変更後:"C:\Program Files\LabelHelper\CSharp_WinAPI_TextShow.exe" "%1"
                        //%1がクォートしてある点に注意、Program Files～はハードリンクで作成
                        //System.Diagnostics.ProcessStartInfo psiLbl = new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\用品管理ラベル出力\label002.exe", "\"" + args[0] + "\"");
                        System.Diagnostics.ProcessStartInfo psiLbl = new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\用品管理ラベル出力\label002.exe", args[0]);
                        System.Diagnostics.Process.Start(psiLbl);
                        //ここから独自の処理を実装していく
                        //MessageBox.Show(System.IO.Path.GetFullPath(args[0]));
                        break;
                    default:
                        //想定外のファイルだった場合は
                        MessageBox.Show("想定外のファイルが引数に渡されました");
                        return;
                }
                if (strExtention == ".txt")
                {
                }
            }
            //Application.Run(new Form1());
        }
    }
}
