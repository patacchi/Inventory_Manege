using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharp_WinAPI_TextShow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (FontFamily item in new InstalledFontCollection().Families) 
            {
                if (item.IsStyleAvailable(FontStyle.Regular))
                {
                    cmbBox_FontNameList.Items.Add(item.Name);
                }
            }
        }
        /// <summary>
        /// コンボボックスをOwnerDrawにしている時に発生、項目を描画するさいに
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbBox_FontNameList_DrawItem(object sender, DrawItemEventArgs e)
        {
            //背景を描画する
            e.DrawBackground();
            ComboBox cmbBox = (ComboBox)sender;
            //項目に表示する文字列を設定
            string strShow = e.Index > -1 ? cmbBox.Items[e.Index].ToString() : cmbBox.Text;
            //使用するフォント
            Font fontDrawing = new Font(strShow, cmbBox.Font.Size);
            //使用するブラシを設定
            Brush brushDrawing = new SolidBrush(e.ForeColor);
            //文字列を描画する
            float flym = (e.Bounds.Height - e.Graphics.MeasureString(strShow, fontDrawing).Height) / 2;
            e.Graphics.DrawString(strShow, fontDrawing, brushDrawing, e.Bounds.X, e.Bounds.Y + flym);
            fontDrawing.Dispose();
            brushDrawing.Dispose();
            //フォーカスを表す四角形を描画
            e.DrawFocusRectangle();
        }
        /// <summary>
        /// コンボボックスでOwnerDrawVariable(行の高さ可変)とした時にのみ発生するイベント
        /// 行の高さを個別に設定できる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbBox_FontNameList_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            ComboBox cmbBox = (ComboBox)sender;
            string strFontName = e.Index > -1 ? cmbBox.Items[e.Index].ToString() : cmbBox.Text;
            //使用するフォント
            Font fontMesure = new Font(strFontName, cmbBox.Font.Size);
            //項目の高さを決定
            e.ItemHeight = (int)e.Graphics.MeasureString(strFontName, fontMesure).Height;
            fontMesure.Dispose();
        }
        private bool SetCmbBox_Size()
        {
            int intSizeMax = 78;
            int intSizeCurrent = 7;
            CmbBox_Size.Items.Clear();
            do
            {
                CmbBox_Size.Items.Add(intSizeCurrent.ToString());
                intSizeCurrent = (int)(intSizeCurrent * 1.5);
            } while (intSizeCurrent <= intSizeMax);
            return false;
        }

        private void CmbBox_FontNameList_SelectionChangeCommitted(object sender, EventArgs e)
        {
            SetCmbBox_Size();
        }

        private void CmbBox_Size_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (cmbBox_FontNameList.SelectedIndex <= 0)
                return;
            if (e.Index <= 0)
                return;
            //背景を描画する
            e.DrawBackground();
            ComboBox cmbBox = (ComboBox)sender;
            //項目に表示する文字列を設定
            string strShow = e.Index > -1 ? cmbBox.Items[e.Index].ToString() : cmbBox.Text;
            //使用するフォント
            Font fontDrawing = new Font(cmbBox_FontNameList.Text,float.Parse(cmbBox.Items[e.Index].ToString()));
            //使用するブラシを設定
            Brush brushDrawing = new SolidBrush(e.ForeColor);
            //文字列を描画する
            float flym = (e.Bounds.Height - e.Graphics.MeasureString(strShow, fontDrawing).Height) / 2;
            e.Graphics.DrawString(strShow, fontDrawing, brushDrawing, e.Bounds.X, e.Bounds.Y + flym);
            fontDrawing.Dispose();
            brushDrawing.Dispose();
        }

        private void CmbBox_Size_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (cmbBox_FontNameList.SelectedIndex <= 0)
                return;
            if (e.Index <= 0)
                return;
            ComboBox cmbBox = (ComboBox)sender;
            //項目に表示する文字列を設定
            string strShow = e.Index > -1 ? cmbBox.Items[e.Index].ToString() : cmbBox.Text;
            //使用するフォント
            Font fontMesure = new Font(cmbBox_FontNameList.Text, float.Parse(cmbBox.Items[e.Index].ToString()));
            //項目の高さを決定
            e.ItemHeight = (int)e.Graphics.MeasureString(strShow, fontMesure).Height;
            e.ItemWidth = (int)e.Graphics.MeasureString(strShow, fontMesure).Width;
            if (cmbBox.Width <= e.ItemWidth)
            {
                cmbBox.Width = e.ItemWidth;
            }

            fontMesure.Dispose();
        }
    }
    
}
