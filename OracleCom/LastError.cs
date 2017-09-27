using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleCom
{
    public class LastError
    {
        #region エラー情報
        private int _lastErrorNumber = 0;
        private string _lastErrorString = "";

        #endregion

        /// <summary>
        /// エラー情報をクリアする
        /// </summary>
        public void LastServerErrReset()
        {
            _lastErrorNumber = 0;
            _lastErrorString = "";

        }

        /// <summary>
        /// エラーコードを取得する
        /// </summary>
        /// <returns></returns>
        public int LastServerErr()
        {
            return _lastErrorNumber;
        }

        /// <summary>
        /// エラーメッセージ（文字列）を取得する
        /// </summary>
        /// <returns></returns>
        public string LastServerErrText()
        {
            return _lastErrorString;
        }

        internal void SetError(int errorNumber,string errorString)
        {
            _lastErrorNumber = errorNumber;
            _lastErrorString = errorString;
        }
    }
}
