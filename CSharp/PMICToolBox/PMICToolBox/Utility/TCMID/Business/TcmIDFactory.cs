using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LIMIT_FILE_TYPE = PmicAutomation.Utility.TCMID.DataStructure.EnumStore.LIMIT_FILE_TYPE;

namespace PmicAutomation.Utility.TCMID.Business
{
    class TcmIDFactory
    {
        private static TcmIDFactory _instance = null;

        private TcmIDFactory() { }

        public static TcmIDFactory GetInstance()
        {
            if (_instance == null)
                _instance = new TcmIDFactory();
            return _instance;
        }

        public TcmIDGenBase GetTcmIDObject(LIMIT_FILE_TYPE type)
        {
            TcmIDGenBase obj = null;
            switch (type)
            {
                case LIMIT_FILE_TYPE.CONTI:
                    obj = new TcmIDGenConti();
                    break;

                case LIMIT_FILE_TYPE.IDS:
                    obj = new TcmIDGenIDS();
                    break;

                case LIMIT_FILE_TYPE.LEAKAGE:
                    obj = new TcmIDGenLeakage();
                    break;

                case LIMIT_FILE_TYPE.OTHERS:
                default:
                    obj = new TcmIDGenOthers();
                    break;
            }
            return obj;
        }
    }
}
