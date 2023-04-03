using FluentMigrator;
using EasyMigrator;
namespace CSharp_DBHandle.CSDB_COMServer.Entity

{
    /// <summary>
    /// Entityで使う定数の定義
    /// </summary>
    static class Const_Entity
    {
        public const string DEFAULT_SETTING_JSON_PATH = @"C:\Users\q3005sbe\AppData\Local\Rep\Inventorymanege\bin\SettingJson\INVGeneral.json";
    }

    [Migration(20230403110002)]
    public class CreateNewTanble : Migration
    {
        public override void Down()
        {
            Delete.Table<T_INV_Label_Temp>();
            Delete.Table("Log");
            // throw new NotImplementedException();
        }

        public override void Up()
        {
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Exists())
            {
                // Create.Table<T_INV_Label_Temp>();
                Create.Table(nameof(T_INV_Label_Temp));
                Create.Columns<T_INV_Label_Temp>();
            }
/*             else
            {
                Alter.Table(nameof(T_INV_Label_Temp))
                    .AlterColumn(nameof(T_INV_Label_Temp.F_Seq))
                        .AsInt32().Identity();
                Alter.Table(nameof(T_INV_Label_Temp))
                    .AlterColumn(nameof(T_INV_Label_Temp.F_INV_Label_Status))
                        .AsString();
                if (!Schema.Table(nameof(T_INV_Label_Temp)).Column("TestAddColumn").Exists())
                {
                    Alter.Table(nameof(T_INV_Label_Temp))
                        .AddColumn("TestAddColumn")
                        .AsInt32()
                        .Nullable();
                }
            } */
            //Logテーブルも作成してみる
            if (!Schema.Table("Log").Exists())
            {
                Create.Table("Log")
                    .WithColumn("ID").AsInt32().Identity()
                    .WithColumn("Text").AsString();
            }
        }
    }
    [Name("T_INV_Label_Temp")]
    public class T_INV_Label_Temp
    {
        public enum enumLabelType
        {
            設定なし = 0,
            直行 = 5,
            出庫 = 6,
            直行_後送 = 7
        }
        /// <summary>
        /// オートインクリメント型のプライマリーキー
        /// </summary>
        /// <value></value>
        [Pk]
        [AutoInc(int.MinValue,1)]
        public int F_Seq{get;set;}
        /// <summary>
        /// ラベルの印刷状態などのフラグを立てたUInt64
        /// </summary>
        /// <value></value>
        [NotNull]
        public Int64 F_INV_Label_Status {get;set;} = 0;
        // public string LabelTempTableName { get; set; } = "";
        /// <summary>
        /// 手配コード文字列長
        /// </summary>
        /// <value></value>
        public Int32 F_INV_Tehaicode_Length { get; set; } = 0;
        /// <summary>
        /// オーダーNo
        /// </summary>
        /// <value></value>
        public string? F_INV_OrderNumber { get; set; } = string.Empty;
        /// <summary>
        /// 識別名 SavePoint
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_Savepoint { get; set; } = string.Empty;
        /// <summary>
        /// フォーム開始時間
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_FormStartTime { get; set; } = string.Empty;
        /// <summary>
        /// 機種
        /// </summary>
        /// <value></value>
        public string? F_INV_Kishu { get; set; } = string.Empty;
        /// <summary>
        /// 棚番(一応ローカルになってるけど・・・ラベルファイルから拾ったシステムのをそのまま入れるかも？)
        /// </summary>
        /// <value></value>
        public string? F_INV_Tana_Local_Text { get; set; } = string.Empty;
        /// <summary>
        /// 手配コード Nullはだめ
        /// </summary>
        /// <value></value>
        
        [Length(30)]
        public string F_INV_Tehai_Code { get; set; } = "";
        /// <summary>
        /// 貯蔵記号 FA BS BL
        /// </summary>
        /// <value></value>
        public string? F_INV_Store_Code { get; set; } = string.Empty;
        /// <summary>
        /// 品名1
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_Name_1 { get; set; } = string.Empty;
        /// <summary>
        /// 品名2 ラベルに出力する時にメインで使用する
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_Name_2 { get; set; } = string.Empty;
        /// <summary>
        /// 備考1
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_Remark_1 { get; set; } = string.Empty;
        /// <summary>
        /// 備考2
        /// </summary>
        /// <value></value>
        public string? F_INV_Label_Remark_2 { get; set; } = string.Empty;
        /// <summary>
        /// 入力日時 ラベルファイルから取得したときは、ラベルファイルのものにする
        /// </summary>
        /// <value></value>
        public string? F_InputDate { get; set; } = string.Empty;
        /// <summary>
        /// 製番
        /// </summary>
        /// <value></value>
        public string? F_INV_Seiban { get; set; } = string.Empty;
        /// <summary>
        /// SBL(SBL→MLNoの順で印刷)
        /// </summary>
        /// <value></value>
        public string? F_INV_SBL { get; set; } = string.Empty;
        /// <summary>
        /// MLNo(SBL→MLNoの順で印刷)
        /// </summary>
        /// <value></value>
        public string? F_INV_ML_No { get; set; } = string.Empty;
        /// <summary>
        /// ラベル種別のコード、別途マスターが必要→Enumで指定した
        /// </summary>
        /// <value></value>
        [DbType(System.Data.DbType.Int32)]
        public enumLabelType F_INV_Label_Type_Code { get; set; } = 0 ;
        /// <summary>
        /// 払出 数量(値操作出来る方)
        /// </summary>
        /// <value></value>
        public Int32? F_INV_Current_Amount {get;set;} = 0;
        /// <summary>
        /// 要求数量、システムで設定された初期数量
        /// </summary>
        /// <value></value>
        public Int32 F_INV_Require_Amount {get;set;} = 0;
    }

}

