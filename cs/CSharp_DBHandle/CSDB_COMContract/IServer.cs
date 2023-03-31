using System;
using System.Runtime.InteropServices;


[ComVisible(true)]
[Guid(ContractGuids.ACCDBServerInterface)]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IAccdbServer
{
    /// <summary>
    /// DBのパスを扱うプロパティ(get,set)
    /// </summary>
    /// <value>DBファイルのフルパスを指定する</value>
    string DBPath { get; set;}
    /// <summary>
    /// 実行するSQL文(get,set)
    /// </summary>
    /// <value>SQL設定</value>
    string SQL{get;set;}
    /// <summary>
    /// DBに接続するための接続文字列(get,set)
    /// </summary>
    /// <value></value>
    string ConnectionString{get;set;}
    /// <summary>
    /// 接続されているかどうか返す(get)
    /// </summary>
    /// <value>接続されていたらTrueを返す</value>
    bool Conected{get;}
    /// <summary>
    /// 実行した結果を格納するJSON(get)
    /// </summary>
    /// <value></value>
    string strResultJSON{get;}
    string DoSQL_With_NO_Transaction
    (
        string strSQL = "",
        string strDBPath= ""
    );
}