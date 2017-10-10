using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Oracle.ManagedDataAccess.Client;

namespace OracleCom
{
    [Guid(DBConnection.ClassId)]
    [ComVisible(true)]
    public class DBConnection : LastError
    {
        public const string ClassId = "EB2B68A6-F341-4BB7-AC6E-7AA8C0E82506";
        OracleConnection _oracleConnection = null;

        readonly string CONNECTION_STRING = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)(HOST={0})(PORT={1}))(CONNECT_DATA=(SERVICE_NAME={2})));User Id={3};Password={4}";

        OracleTransaction _Trans = null;

        /// <summary>
        /// データベース接続
        /// </summary>
        /// <param name="HostName"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="PortNumber"></param>
        /// <param name="ServiceName"></param>
        /// <returns></returns>
        public DBCommand OpenDatabase(string HostName, string UserName,string Password,int PortNumber,string ServiceName)
        {
            if (_oracleConnection == null)
            {
                _oracleConnection = new OracleConnection();

                var sb = new StringBuilder();
                sb.AppendFormat(CONNECTION_STRING, HostName, PortNumber, ServiceName, UserName, Password);
                _oracleConnection.ConnectionString = sb.ToString();

                try
                {
                    _oracleConnection.Open();
                }
                catch (Exception ex)
                {
                    SetError(-1, ex.Message);
                    return null;
                }
            }

            LastServerErrReset();
            return new DBCommand(_oracleConnection);
        }

        /// <summary>
        /// データベース切断
        /// </summary>
        public void Close()
        {
            if (_oracleConnection != null)
            {
                _oracleConnection.Close();
                _oracleConnection = null;
            }
        }

        /// <summary>
        /// トランザクションを開始する
        /// </summary>
        public void BeginTrans()
        {
            if (_Trans != null)
            {
                SetError(-1, "Transactionは開始中です");
                return;
            }

            _Trans = _oracleConnection.BeginTransaction();
            LastServerErrReset();
        }

        /// <summary>
        /// トランザクションをコミットする
        /// </summary>
        public void CommitTrans()
        {
            if (_Trans == null)
            {
                SetError(-1, "Transactionは開始されていません");
            }

            try
            {
                _Trans.Commit();
            }catch(Exception ex)
            {
                SetError(-1, ex.Message);
            }
        }

        /// <summary>
        /// トランザクションをロールバックする
        /// </summary>
        public void Rollback()
        {
            if (_Trans == null)
            {
                SetError(-1, "Transactionは開始されていません");
            }
            try
            {
                _Trans.Rollback();
            }
            catch (Exception ex)
            {
                SetError(-1, ex.Message);
            }

        }
    }
}
