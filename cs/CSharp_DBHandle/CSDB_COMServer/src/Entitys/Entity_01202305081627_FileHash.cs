using FluentMigrator;
using CSDB_COMServer.Utility;

namespace CSDB_COMServer.Entitys
{
    /// <summary>
    /// T_INV_Label_Temp に ファイルハッシュのコラムを追加する
    /// </summary>
    [EnforceMigrationNumber(01,2023,05,08,16,27,"Daisuke Oota")]
    public class T_INV_Label_Temp_FileHash : Migration
    {
        public override void Down()
        {
            if(Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_FileHash)).Exists())
            {
                Delete.Column(nameof(T_INV_Label_Temp.F_FileHash)).FromTable(nameof(T_INV_Label_Temp));
            }
        }

        public override void Up()
        {
            if(!Schema.Table(nameof(T_INV_Label_Temp)).Column(nameof(T_INV_Label_Temp.F_FileHash)).Exists())
            {
                Create.Column(nameof(T_INV_Label_Temp.F_FileHash)).OnTable(nameof(T_INV_Label_Temp)).AsString(64).Nullable();
            }
        }
    }
    public partial class T_INV_Label_Temp
    {
        /// <summary>
        /// 格納するファイルのハッシュ
        /// </summary>
        /// <value>SHA512で求めて、16進表記の文字列とする(文字列数 128)</value>
        public string? F_FileHash {get;set;} = string.Empty;
    }
}
