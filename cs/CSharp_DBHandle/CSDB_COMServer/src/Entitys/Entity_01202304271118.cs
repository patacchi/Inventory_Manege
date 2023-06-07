using FluentMigrator;
using EasyMigrator;
using CSDB_COMServer.Utility;
namespace CSDB_COMServer.Entitys
{

    [EnforceMigrationNumber(01,2023,04,27,11,18,"Daisuke Oota")]
    public class CreateNewTableOne : Migration
    {
        public override void Down()
        {
            // Delete.Table<T_INV_Label_Temp>();
            if (Schema.Table(nameof(T_INV_Label_Temp)).Exists())
            {
                Delete.Table(nameof(T_INV_Label_Temp));
            }
            if (Schema.Table("Log").Exists())
            {
                Delete.Table("Log");
            }
            // throw new NotImplementedException();
        }

        public override void Up()
        {
            bool isNewTable = false;
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Exists())
            {
                //NewTableフラグを上げる
                isNewTable = true;
                // Create.Table<T_INV_Label_Temp>();
                //テーブルが存在しなかった場合は、とりあえるキーになるF_Seqのみのテーブルを作成する
                Create.Table(nameof(T_INV_Label_Temp))
                .WithColumn(nameof(T_INV_Label_Temp.F_Seq)).AsInt32().NotNullable().PrimaryKey().Identity();
            }
            //その他のフィールドは、それぞれ存在の有無を確認しながら追加する
            //F_Seqについては、テーブル新規作成時にはまだフィールド存在していないので、新規テーブルフラグと合わせて判断する
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_Seq)).Exists() && !isNewTable)
            {
                //F_seqが無かった場合(外部でテーブル新規作成？)
                //まずはフィールドを作成(オートインクリメント)
                Create.Column(nameof(T_INV_Label_Temp.F_Seq)).OnTable(nameof(T_INV_Label_Temp))
                .AsInt32().NotNullable().Identity();
                Create.PrimaryKey().OnTable(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_Seq));
            }
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_InputDate)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_InputDate)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Tehai_Code)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Tehai_Code)).OnTable(nameof(T_INV_Label_Temp)).AsString(30);
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Tehaicode_Length)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Tehaicode_Length)).OnTable(nameof(T_INV_Label_Temp)).AsInt32();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Current_Amount)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Current_Amount)).OnTable(nameof(T_INV_Label_Temp)).AsInt32().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Kishu)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Kishu)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_FormStartTime)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_FormStartTime)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Name_1)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Name_1)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Name_2)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Name_2)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Remark_1)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Remark_1)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Remark_2)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Remark_2)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Savepoint)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Savepoint)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Status)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Status)).OnTable(nameof(T_INV_Label_Temp)).AsInt64().WithDefaultValue(0);
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Label_Type_Code)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Label_Type_Code)).OnTable(nameof(T_INV_Label_Temp)).AsInt32().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_ML_No)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_ML_No)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_OrderNumber)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_OrderNumber)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Require_Amount)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Require_Amount)).OnTable(nameof(T_INV_Label_Temp)).AsInt32().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_SBL)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_SBL)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Seiban)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Seiban)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Store_Code)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Store_Code)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
            if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Tana_Local_Text)).Exists())
            Create.Column(nameof(T_INV_Label_Temp.F_INV_Tana_Local_Text)).OnTable(nameof(T_INV_Label_Temp)).AsString().Nullable();
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
    public partial class T_INV_Label_Temp :IEquatable<T_INV_Label_Temp>
    {
        public enum enumLabelType
        {
            設定なし = 0,
            直行 = 5,
            出庫 = 6,
            直行_後送 = 7
        }

        #region ColumnField
        /// <summary>
        /// オートインクリメント型のプライマリーキー
        /// </summary>
        /// <value></value>
        [Pk]
        [AutoInc(int.MinValue,1)]
        [NotIncludingValueList]
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
        public Int32? F_INV_Tehaicode_Length { get; set; } = 0;
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
        [NotNull]
        // public string F_INV_Tehai_Code { get; set; } = " ";
        public string F_INV_Tehai_Code { get; set; } = string.Empty;
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
        // public enumLabelType? F_INV_Label_Type_Code { get; set; } = 0 ;
        public int? F_INV_Label_Type_Code { get; set; } = 0 ;
        /// <summary>
        /// 払出 数量(値操作出来る方)
        /// </summary>
        /// <value></value>
        public Int32? F_INV_Current_Amount {get;set;} = 0;
        /// <summary>
        /// 要求数量、システムで設定された初期数量
        /// </summary>
        /// <value></value>
        public Int32? F_INV_Require_Amount {get;set;} = 0;
        #endregion 

        public bool Equals(T_INV_Label_Temp? other)
        {
            if (other == null)
            {
                return false;
            }
            //引数を比較対象クラスでキャストする
            T_INV_Label_Temp otherclass = (T_INV_Label_Temp)other;
            //Entityクラス同士が同一である条件
            //FileHash,
            return (this.F_FileHash == otherclass.F_FileHash && this.F_InputDate == otherclass.F_InputDate);
        }

    }

}

