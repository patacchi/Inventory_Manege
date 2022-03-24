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
                string strExtention = System.IO.Path.GetExtension(args[0]);
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
                }
                if (strExtention == ".txt")
                {
                }
            }
            Application.Run(new Form1());
        }
    }
}
