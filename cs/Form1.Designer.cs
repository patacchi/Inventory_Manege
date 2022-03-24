namespace CSharp_WinAPI_TextShow
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_ShowTextWindow = new System.Windows.Forms.Button();
            this.cmbBox_FontNameList = new System.Windows.Forms.ComboBox();
            this.CmbBox_Size = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btn_ShowTextWindow
            // 
            this.btn_ShowTextWindow.Location = new System.Drawing.Point(163, 64);
            this.btn_ShowTextWindow.Name = "btn_ShowTextWindow";
            this.btn_ShowTextWindow.Size = new System.Drawing.Size(168, 25);
            this.btn_ShowTextWindow.TabIndex = 0;
            this.btn_ShowTextWindow.Text = "上記内容のウィンドウ表示";
            this.btn_ShowTextWindow.UseVisualStyleBackColor = true;
            // 
            // cmbBox_FontNameList
            // 
            this.cmbBox_FontNameList.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.cmbBox_FontNameList.FormattingEnabled = true;
            this.cmbBox_FontNameList.Location = new System.Drawing.Point(37, 17);
            this.cmbBox_FontNameList.Name = "cmbBox_FontNameList";
            this.cmbBox_FontNameList.Size = new System.Drawing.Size(187, 20);
            this.cmbBox_FontNameList.TabIndex = 1;
            this.cmbBox_FontNameList.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.CmbBox_FontNameList_DrawItem);
            this.cmbBox_FontNameList.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.CmbBox_FontNameList_MeasureItem);
            this.cmbBox_FontNameList.SelectionChangeCommitted += new System.EventHandler(this.CmbBox_FontNameList_SelectionChangeCommitted);
            // 
            // CmbBox_Size
            // 
            this.CmbBox_Size.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.CmbBox_Size.FormattingEnabled = true;
            this.CmbBox_Size.Location = new System.Drawing.Point(260, 20);
            this.CmbBox_Size.Name = "CmbBox_Size";
            this.CmbBox_Size.Size = new System.Drawing.Size(113, 20);
            this.CmbBox_Size.TabIndex = 2;
            this.CmbBox_Size.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.CmbBox_Size_DrawItem);
            this.CmbBox_Size.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.CmbBox_Size_MeasureItem);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(489, 101);
            this.Controls.Add(this.CmbBox_Size);
            this.Controls.Add(this.cmbBox_FontNameList);
            this.Controls.Add(this.btn_ShowTextWindow);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_ShowTextWindow;
        private System.Windows.Forms.ComboBox cmbBox_FontNameList;
        private System.Windows.Forms.ComboBox CmbBox_Size;
    }
}

