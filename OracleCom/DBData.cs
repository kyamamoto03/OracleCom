using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Collections.Specialized;

namespace OracleCom
{

    [ComVisible(true)]
    public class DBData
    {
        List<OrderedDictionary> datas = new List<OrderedDictionary>();
        int index = 0;
        bool EofFlag = true;

        public DBData()
        {

        }

        /// <summary>
        /// データを追加する
        /// </summary>
        /// <param name="data"></param>
        internal void Add(OrderedDictionary data)
        {
            datas.Add(data);
            EofFlag = false;
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

            EofFlag = true;
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
        public object this[int Key]
        {
            get
            {
                return datas[index][Key];
            }
        }

        public void Close()
        {
            //Nothing To Do
        }

        /// <summary>
        /// 項目(列)数を返す
        /// </summary>
        /// <returns></returns>
        public int ColumCount()
        {
            if (datas.Count > 0)
            {
                return datas[0].Count;
            }
            else
            {
                return 0;
            }

        }
        /// <summary>
        /// FieldDataを返す
        /// データが０件の場合はエラー
        /// </summary>
        /// <returns></returns>
        public FieldData Fields
        {
            get
            {
                return new FieldData(datas[index]);
            }
        }
        /// <summary>
        /// データ件数を返す
        /// </summary>
        /// <returns></returns>
        public int RecordCount()
        {
            return datas.Count;
        }

        /// <summary>
        /// EOFを返す
        /// </summary>
        /// <returns></returns>
        public bool EOF()
        {
            return EofFlag;

        }

        public class FieldData
        {
            OrderedDictionary fieldData;

            public FieldData(OrderedDictionary data)
            {
                this.fieldData = data;
            }
            public int Count
            {
                get
                {
                    return this.fieldData.Count;
                }
            }
        }
    }
}
