using System;
using System.Runtime.InteropServices;
using Dapper;

namespace CSharp_DBHandle.CSDB_COMServer
{
    [ComVisible(true)]
    [Guid(ContractGuids.ACCDBServerClass)]
    [ProgId("CSharp.ACCDB.COMServer")]
    public class AccdbHandle : IAccdbServer
    {
        public string DBPath { get 
        {
            return "工事中";
             throw new NotImplementedException();
        } 
         set => throw new NotImplementedException(); }
        public string SQL { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string ConnectionString { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public bool Conected => throw new NotImplementedException();

        public string strResultJSON => throw new NotImplementedException();

        public string DoSQL_With_NO_Transaction(string strSQL = "", string strDBPath = "")
        {
            throw new NotImplementedException();
        }
/*         public List<CSDB_COMServer.Entity.T_INV_Label_Temp> labelTable_Data
        {
            get 
            {
            
                .Entity.T_INV_Label_Temp labelresult;
                labelresult = new Entity.T_INV_Label_Temp();
                return labelresult;
            }
        } */
    }
}