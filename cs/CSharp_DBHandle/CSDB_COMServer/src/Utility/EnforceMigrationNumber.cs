using System;


namespace CSDB_COMServer.Utility
{
    /// <summary>
    /// Migration バージョン番号生成クラス
    /// </summary>
    public class EnforceMigrationNumber : FluentMigrator.MigrationAttribute
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="branchnumber_">(num)****</param>
        /// <param name="year_">yyyy</param>
        /// <param name="month_">MM</param>
        /// <param name="day_">dd</param>
        /// <param name="hour_">hh</param>
        /// <param name="minute_">mm</param>
        /// <param name="author_">コメント用</param>
        public EnforceMigrationNumber(int branchnumber_,int year_,int month_,int day_,int hour_,int minute_,string author_)
        :base (CaluculateValue(branchnumber_,year_,month_,day_,hour_,minute_))
        {
            this.Author = author_;
        }
        public string Author {get;private set;}
        private static long CaluculateValue(int branchnumber,int year,int month,int day,int hour,int minute)
        {
            //(Branchnumber)202304271123
            return branchnumber * 1000000000000L + year * 100000000L + month * 1000000L + day * 10000L + hour * 100L + minute;
        }
    }
}
