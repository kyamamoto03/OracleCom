using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OracleCom
{

    [ComVisible(true)]
    public class DBData
    {
        List<Dictionary<string, object>> datas = new List<Dictionary<string, object>>();
        int index = 0;

        public DBData()
        {

        }

        /// <summary>
        /// データを追加する
        /// </summary>
        /// <param name="DBData"></param>
        internal void Add(Dictionary<string, object> DBData)
        {
            datas.Add(DBData);
        }

        /// <summary>
        /// 次の行へ移動する
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            index++;
            if (datas.Count > index)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="Key"></param>
        /// <returns></returns>
        public object this[string Key]
        {
            get
            {
                return datas[index][Key.ToUpper()];
            }
        }

        /// <summary>
        /// データ件数を返す
        /// </summary>
        /// <returns></returns>
        public int Count()
        {
            return datas.Count;
        }
    }
}
