using System;

namespace CSDB_COMServer.Utility
{
    /// <summary>
    /// Entityに付ける属性の定義を行うクラス
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class NotIncludingValueListAttribute :System.Attribute
    {
        private bool _isNotInclude;
        public NotIncludingValueListAttribute()
        {
            _isNotInclude = true;
        }
        public NotIncludingValueListAttribute(bool? isNotInclude = true)
        {
            _isNotInclude = (bool)isNotInclude!;
        }
    }
}
