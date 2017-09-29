using System.Collections.Generic;
using System.Runtime.InteropServices;
using Oracle.ManagedDataAccess.Client;

namespace OracleCom
{
    [ComVisible(true)]
    public class DBCommand : LastError
    {
        OracleConnection _oracleConnection;

        public DBCommand(OracleConnection conn)
        {
            _oracleConnection = conn;
        }

        public DBData CreateDynaset(string SQL,int Param)
        {
            return CreateDynaset(SQL);
        }

        public DBData CreateDynaset(string SQL)
        {
            DBData datas = new DBData();


            if (_oracleConnection != null)
            {
                using (OracleCommand cmd = new OracleCommand(SQL, _oracleConnection))
                {
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Dictionary<string, object> data = new Dictionary<string, object>();
                            for (int i = 0; i < dr.VisibleFieldCount; i++)
                            {
                                data.Add(dr.GetName(i).ToUpper(), dr.GetValue(i));
                            }
                            datas.Add(data);

                        }
                    }

                }
            }

            return datas;
        }

        /// <summary>
        /// SQLを実行する
        /// </summary>
        /// <param name="SQL"></param>
        public int ExecuteSQL(string SQL)
        {
            if (_oracleConnection != null)
            {
                using (OracleCommand cmd = new OracleCommand(SQL, _oracleConnection))
                {
                    return cmd.ExecuteNonQuery();
                }
            }
            else
            {
                SetError(-1, "データベースの接続がありません");
            }
            return 0;
        }
    }
}
