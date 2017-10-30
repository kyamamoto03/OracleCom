using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Collections.Specialized;
using System.Linq;

namespace OracleCom
{

    [ComVisible(true)]
    public class DBData
    {
        List<DBDataDetail> datas = new List<DBDataDetail>();
        DBHeader _dbHeader;
        int index = 0;
        bool EofFlag = true;

        public DBData()
        {
        }

        public void SetHeader(DBHeader dbHeader)
        {
            _dbHeader = dbHeader;
        }

        /// <summary>
        /// データを追加する
        /// </summary>
        /// <param name="data"></param>
        internal void Add(DBDataDetail data)
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
                return datas[index][_dbHeader[Key.ToUpper()]];
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
                return _dbHeader.Count;
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
        public  int Count
        {
            get
            {
                return datas.Count;
            }
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
            DBDataDetail fieldData;

            public FieldData(DBDataDetail data)
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
            public object this[int index]
            {
                get
                {
                    return fieldData[index];
                }
            }
        }
    }

    /// <summary>
    /// 行データを管理するクラス
    /// </summary>
    public class DBDataDetail
    {
        List<object> _datas = new List<object>();

        public void Add(object obj)
        {
            _datas.Add(obj);
        }

        public object this[int index]
        {
            get
            {
                return _datas[index];
            }
        }

        public int Count
        {
            get
            {
                return _datas.Count;
            }
        }
    }

    /// <summary>
    /// カラム名を管理するクラス
    /// </summary>
    public class DBHeader
    {
        OrderedDictionary _ColumnNames = new OrderedDictionary();
        int _index = 0;

        public void AddColumnName(string columnName)
        {
            _ColumnNames.Add(columnName,_index++);
        }

        public int Count
        {
            get
            {
                return _ColumnNames.Count;
            }
        }
        public int this[string columnName]
        {
            get
            {
                int index;
                index = (int)_ColumnNames[columnName];
                return index;
            }
        }
    }
}
