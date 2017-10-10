using System.Collections.Generic;
using System.Runtime.InteropServices;
using Oracle.ManagedDataAccess.Client;
using System.Collections.Specialized;
using System;

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
            try
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
                                OrderedDictionary data = new OrderedDictionary();
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
            }catch(Exception ex)
            {
                SetError(ex.HResult, ex.Message);
                return null;
            }
        }

        /// <summary>
        /// SQLを実行する
        /// </summary>
        /// <param name="SQL"></param>
        public int ExecuteSQL(string SQL)
        {
            try
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
            }catch(Exception ex)
            {
                SetError(ex.HResult, ex.Message);
                return -1;
            }
        }
    }
}
