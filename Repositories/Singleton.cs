using System;
using System.Data.SqlClient;

namespace EindOpdrachtLeandro.Repositories
{   //Sealed zorgt ervoor dat je niet kan overerven
    public sealed class Singleton : IDisposable
    {
        //This implementation is thread-safe.
        //In the following code, the thread is locked on a shared object and checks whether an instance has been created or not.
        //This takes care of the memory barrier issue and ensures that only one thread will create an instance.
        //For example: Since only one thread can be in that part of the code at a time, by the time the second thread enters it, the first thread will have created the instance, so the expression will evaluate to false.
        //The biggest problem with this is performance; performance suffers since a lock is required every time an instance is requested.

        Singleton()
        {

        }
        private static readonly object padlock = new object();
        private static Singleton dbInstance = null;
        private const string sConnectieString =
             @"Data source =DESKTOP-PRJJRNL\SQLVIVES; Initial Catalog=EindwerkLeandro; Integrated Security=True";
        // @"Data source = ISABELLE-HPPRO\VIVES; Initial Catalog=Bookstore; Integrated Security=True";
        private readonly SqlConnection conn = new SqlConnection(sConnectieString);


        public static Singleton Instance
        {
            get
            {
                lock (padlock)
                {
                    if (dbInstance == null)
                    {
                        dbInstance = new Singleton();
                    }
                    return dbInstance;
                }
            }
        }


        public SqlConnection GetDBConnection()
        {
            try
            {
                conn.Open();
                Console.WriteLine("Connected");
            }
            catch (SqlException e)
            {
                Console.WriteLine("Not connected : " + e.ToString());
                Console.ReadLine();
            }
            finally
            {
                Console.WriteLine("End..");
                // Console.WriteLine("Not connected : " + e.ToString());
                Console.ReadLine();
            }
            Console.ReadLine();
            return conn;
        }

        //sluiten van verbinding --> afsluiten van programma
        //gebruik van Dispose-methode die vastgelegd wordt in de IDisposable-interface
        public void Dispose()
        {
            if (sConnectieString != null)
            {
                conn.Dispose();
            }
        }
    }
}