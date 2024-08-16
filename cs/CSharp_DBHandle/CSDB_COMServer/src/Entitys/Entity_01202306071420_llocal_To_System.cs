using FluentMigrator;
using CSDB_COMServer.Utility;

namespace CSDB_COMServer.Entitys
{
    /// <summary>
    /// ラベルから取得するTanaをSystemとして扱うため
    /// 新規に_Tana_System_フィールドを作成する
    /// </summary>
    [EnforceMigrationNumber(01,2023,06,07,14,20,"Daisuke Oota")]
    public class T_INV_Temp_Tana_Local_to_System : Migration
    {
        public override void Down()
        {
            //System_textが存在する場合は削除する
            if (Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Tana_System_Text)).Exists())
            {
                Delete.Column(nameof(T_INV_Label_Temp.F_INV_Tana_System_Text)).FromTable(nameof(T_INV_Label_Temp));
            }
        }
        public override void Up()
        {
            //System_textが存在しない場合、フィールド追加
             if (!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_INV_Tana_System_Text)).Exists())
             {
                Create.Column(nameof(T_INV_Label_Temp.F_INV_Tana_System_Text)).OnTable(nameof(T_INV_Label_Temp)).AsString(10).Nullable();
             }
        }
    }
    public partial class T_INV_Label_Temp
    {
        /// <summary>
        /// Systemの棚番
        /// </summary>
        /// <value></value>
        public string? F_INV_Tana_System_Text{get;set;} = string.Empty;
    }
}
