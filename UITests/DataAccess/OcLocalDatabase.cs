using System.IO;
using System.Threading;
using LiteDB;

namespace UITests.DataAccess
{
    public class OcLocalDatabase
    {
        public static void AddFavoritesToLiteDb(string dbFilePath)
        {
            var bson = JsonSerializer.DeserializeArray(Resources.FavoriteMatters);
            var dbFilePathInfo = new DirectoryInfo(dbFilePath);

            var connectionString = $"filename={dbFilePathInfo};async=true;";

            if (!WaitForLiteDbCreation(dbFilePathInfo.ToString())) return;

            using (var db = new LiteDatabase(connectionString))
            {
                var collection = db.GetCollection("faves");

                foreach (var bsonValue in bson)
                {
                    collection.Insert(bsonValue.AsDocument);
                }
            }
        }

        private static bool WaitForLiteDbCreation(string databaseFile)
        {
            const int timeout = 20000;
            const int step = 500;
            var elapsed = 0;

            while (!File.Exists(databaseFile))
            {
                Thread.Sleep(step);
                elapsed += step;

                if (elapsed > timeout) throw new FileNotFoundException();
            }

            return true;
        }
    }
}