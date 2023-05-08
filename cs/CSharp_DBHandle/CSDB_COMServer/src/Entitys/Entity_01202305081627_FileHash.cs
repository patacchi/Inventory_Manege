using FluentMigrator;
using CSDB_COMServer.Utility;
using System.ComponentModel.DataAnnotations;

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
            throw new NotImplementedException();
        }

        public override void Up()
        {
            throw new NotImplementedException();
        }
    }
    public partial class T_INV_Label_Temp
    {
        /// <summary>
        /// 格納するファイルのハッシュ SHA512で求める
        /// </summary>
        /// <value>16進表記の文字列とする</value>
        public string? F_FileHash {get;set;} = string.Empty;
    }
}
