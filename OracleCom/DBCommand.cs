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
        OracleCommand _oracleCommand;

        public DBCommand(OracleConnection conn)
        {
            _oracleConnection = conn;
            _oracleCommand = new OracleCommand();
            _oracleCommand.Connection = conn;
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

                bool FirstTime = true;
                if (_oracleConnection != null)
                {
                    _oracleCommand.CommandText = SQL;
                    using (OracleDataReader dr = _oracleCommand.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            if (FirstTime)
                            {
                                DBHeader header = new DBHeader();
                                for (int i = 0; i < dr.VisibleFieldCount; i++)
                                {
                                    header.AddColumnName(dr.GetName(i).ToUpper());
                                }

                                datas.SetHeader(header);
                                FirstTime = false;
                            }

                            DBDataDetail data = new DBDataDetail();
                            for (int i = 0; i < dr.VisibleFieldCount; i++)
                            {
                                data.Add(dr.GetValue(i));
                            }
                            datas.Add(data);

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

        public DBParameters Parameters
        {
            get
            {
                return new DBParameters(_oracleCommand);
            }
        }


    }
    /// <summary>
    /// パラメータのラッピングクラス
    /// </summary>
    [ComVisible(true)]
    public class DBParameters
    {
        public enum DB_TYPE
        {
            ORATYPE_NUMBER = 2,
            ORATYPE_CHAR = 96
        }
        OracleCommand _cmd;
        public DBParameters(OracleCommand cmd)
        {
            _cmd = cmd;
        }

        public void Add(string ParameterName, object value,int DbType)
        {
            switch(DbType)
            {
                case 96:
                    _cmd.Parameters.Add(new OracleParameter(ParameterName, value));
                    break;
                case 2://NUMBER
                    _cmd.Parameters.Add(new OracleParameter(ParameterName,OracleDbType.Int32, value,System.Data.ParameterDirection.InputOutput));
                    break;
                default:
                    throw new Exception("対応していないデータ型です");
            }
        }

        public void Remove(string ParameterName)
        {
            _cmd.Parameters.RemoveAt(ParameterName);
        }

        public object this[string ParameterName]
        {
            get
            {
                if(_cmd.Parameters[ParameterName].OracleDbType == OracleDbType.Int32)
                {
                    return Convert.ToInt32((decimal)(Oracle.ManagedDataAccess.Types.OracleDecimal)(_cmd.Parameters[ParameterName].Value));
                }
                return _cmd.Parameters[ParameterName].Value;
            }

        }

        public object this[int ParameterIndex]
        {
            get
            {
                return _cmd.Parameters[ParameterIndex].Value;
            }

        }
    }
}
