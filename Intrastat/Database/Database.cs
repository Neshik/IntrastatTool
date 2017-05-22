using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Database
{
    static class Database
    {
        private static String databaseDSN = "DSN=Aqualight";

        public static OdbcConnection getDatabase()
        {
            OdbcConnection DbConnection = null;
            try
            {
                DbConnection = new OdbcConnection(databaseDSN);
            }
            catch
            {
                Console.WriteLine("Fehler bei der Datenbankverbindung! Ausgeführt wurde: " + databaseDSN);
            }
            return DbConnection;
        }
    }
}
