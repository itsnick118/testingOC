namespace MockPassport.Mappings.Strings
{
    public class Endpoint
    {
        public const string StaticResourcesRegex = @"\/application\/office\/[v.0-9]+\/officeui\/";

        public const string Passport = "/";

        public const string FetchList = "/enduser/listscreens/fetchList.do";

        public const string OcManifest = "/enduser/office/manifest/ocmanifest.do";

        public const string ListScreenShowJson = "/enduser/listscreens/show.json";

        public const string Index = "/enduser/index.do";

        public const string InitializeSessionParam = 
            "/enduser/office/companion/initializeSessionParam.do";

        public const string ItemScreenShowJson = "/enduser/itemscreens/show.json";

        public const string GlowRootEnabled = "/enduser/index.do&--glowroot-eum";

        public const string GlowRootDisabled = "/enduser/index.do&--glowroot-eum-nf";

        public const string GetUniqueToken = "/enduser/office/companion/getUniqueToken.do";

        public static string GetEntity(string entity)
        {
            return $"{EntityBase}{entity}";
        }

        public static string GetEntity(string entity, int id)
        {
            return $"{EntityBase}{entity}/{id}";
        }

        public static string GetEntity(string entity, string endpoint)
        {
            return $"{EntityBase}{entity}{endpoint}";
        }

        private const string EntityBase = "/datacert/api/entity/";
    }
}