using System;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace AutoTestSystem
{
    public class SqlManager
    {
        private readonly string _sqlDb;
        private SQLiteDataAdapter _sqlAdapter;
        private SQLiteCommand _sqlCmd;
        private SQLiteConnection _sqlCon;
        public string RecordTableName = "Status";

        public SqlManager(string dbName)
        {
            _sqlDb = dbName;
            InitSql();
        }

        private void InitSql()
        {
            if (!File.Exists(_sqlDb) && _sqlDb != "N/A")
            {
                SqLiteCon();

                CreateDbTable();

                SqLiteClose();
            }
        }

        private void SqLiteCon()
        {
            try
            {
                if (_sqlCon == null)
                {
                    _sqlCon = new SQLiteConnection(string.Format("Data Source={0};Compress=True", _sqlDb));
                    _sqlCon.Open();
                    _sqlCmd = new SQLiteCommand(_sqlCon);
                    _sqlAdapter = new SQLiteDataAdapter(_sqlCmd);
                    _sqlAdapter.AcceptChangesDuringUpdate = true;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void SqLiteClose()
        {
            if (_sqlCon != null)
            {
                _sqlCon.Close();
                _sqlCmd = null;
                _sqlAdapter = null;
                _sqlCon = null;
                GC.Collect();
            }
        }

        private void CreateDbTable()
        {
            string[] arr =
            {
                "\"Tester\" string",
                "\"SettingIni\" string",
                "\"TurnOn\" bool",
                "\"Job\" string",
                "\"TestProgram\" string",
                "\"PatternFolder\" string",
                "\"PatternSync\" string",
                "\"LotId\" string",
                "\"WaferId\" string",
                "\"SetXy\" string",
                "\"EnableWords\" string",
                "\"DoAll\" bool"
            };
            _sqlCmd.CommandText = "create table \"" + RecordTableName + "\" (" + string.Join(",", arr) +
                                  ", PRIMARY KEY (Tester,SettingIni))";
            _sqlCmd.ExecuteNonQuery();
        }
    }
}