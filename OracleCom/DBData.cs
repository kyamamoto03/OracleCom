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
        int _columnCount;
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
        /// 前の行へ移動する
        /// </summary>
        /// <returns></returns>
        public bool MovePrevious()
        {
            index--;
            if (0 <= index)
            {
                return true;
            }

            return false;
        }
        /// <summary>
        /// 最初の行へ移動する
        /// </summary>
        /// <returns></returns>
        public void MoveFirst()
        {
            index = 0;
        }
        /// <summary>
        /// 最後の行へ移動する
        /// </summary>
        /// <returns></returns>
        public void MoveLast()
        {
            index = datas.Count - 1;
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
            return _columnCount;
        }
        public int ColumnCount
        {
            get { return _columnCount; }
            set { _columnCount = value; }
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
                if (datas.Count > 0)
                {
                    var f = new FieldData(datas[index]);
                    f.SetHeader(_dbHeader);
                    f.Count = ColumnCount;
                    return f;
                }
                else
                {
                    var f = new FieldData(new DBDataDetail());
                    f.SetHeader(_dbHeader);
                    f.Count = ColumnCount;
                    return f;
                }
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
            //return EofFlag;
            if (index >= datas.Count)
            {
                return true;
            }
            else
            {
                return false;
            }


        }
        /// <summary>
        /// BOFを返す
        /// </summary>
        /// <returns></returns>
        public bool BOF()
        {
            if(index < 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public class FieldData
        {
            DBDataDetail fieldData;
            DBHeader _dbHeader;
            int _columnCount;

            public FieldData(DBDataDetail data)
            {
                this.fieldData = data;
            }

            public void SetHeader(DBHeader dbHeader)
            {
                _dbHeader = dbHeader;
            }

            public int Count
            {
                get
                {
                    return _columnCount;
                }
                internal set
                {
                    _columnCount = value;
                }
            }
            public object this[int index]
            {
                get
                {
                    return fieldData[index];
                }
            }
            public object this[string Key]
            {
                get
                {
                    return fieldData[_dbHeader[Key.ToUpper()]];
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
        int _count;

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
                return _count;
            }
            set
            {
                _count = value;
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
