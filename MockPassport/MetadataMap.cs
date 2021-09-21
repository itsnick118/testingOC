using System.Collections.Generic;
using System.IO;
using System.Xml;
using MockPassport.Mappings.Metadata;

namespace MockPassport
{
    public class MetadataMap
    {
        public IDictionary<string, string> EntityHrefs { get; set; }

        public const string ByIdSuffix = "_by_id";
        public const string ByNameSuffix = "_by_name";

        public MetadataMap(IEnvironment environment)
        {
            EntityHrefs = new Dictionary<string, string>();
            foreach (var entry in EntityNames.MetaDataFiles)
            {
                var file = Path.Combine(environment.BaseFilePath.FullName, "Responses", entry.Value + ByNameSuffix);

                if (!File.Exists(file)) continue;

                var parsedResult = new XmlDocument();
                parsedResult.Load(file);

                if (parsedResult.DocumentElement == null) continue;

                foreach (XmlElement entityNameNode in parsedResult.DocumentElement)
                {
                    EntityHrefs[entityNameNode.InnerText] = entityNameNode.Attributes["href"].Value;
                }
            }
        }
    }
}
