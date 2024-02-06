using ScriptForBaceNetFramework.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptForBaceNetFramework.Class
{
    internal class GetDB
    {
        private GetDB() { }

        private static GetDB instance;

        private BasketballSystemEntities _db;

        public static GetDB GetInstance()
        {
            if (instance == null)
            {
                instance = new GetDB ();
            }
            return instance;
        }

        public BasketballSystemEntities DB()
        {
            if (_db == null)
            {
                _db = new BasketballSystemEntities();
            }
            return _db;
        }
    }
}
