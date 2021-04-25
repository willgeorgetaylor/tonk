using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DotLiquid;
using DotLiquid.Tags;
using Newtonsoft.Json.Linq;
using Syncfusion.DocIO.DLS;

namespace Tonk.DOM
{
    internal class Parser
    {


        public static string Parse(Document doc)
        {
            var array = new JArray();
            var container = new JObject(new JProperty("fields", array));

            foreach (var c in ParseRecursive(doc, array))
            {

            }

            return container.ToString();
        }


        private static IEnumerable<IEntity> ParseRecursive(Tag parentTag, JArray parentObject)
        {
            foreach (var childEntity in parentTag.NodeList)
            {
                if (childEntity is string)
                    continue;

                var currentNode = new JObject();
                parentObject.Add(currentNode);

                currentNode.Add(new JProperty("type", childEntity.ToString()));

                switch (childEntity)
                {
                    case Tag tonkTag:
                        var nodeHolder = new JArray();
                        currentNode.Add(new JProperty("name", tonkTag.Name));

                        if (tonkTag.NodeList == null || tonkTag.NodeList.Count <= 0) continue;
                        
                        foreach (var grandchildEntity in ParseRecursive(tonkTag, nodeHolder))
                        {
                            yield return grandchildEntity;
                        }

                        if (nodeHolder.Count > 0)
                            currentNode.Add(new JProperty("nodes", nodeHolder));
                        break;
                    case Variable tonkVar:
                        currentNode.Add(new JProperty("variable", tonkVar.Name));
                        var filters = JToken.FromObject(from f in tonkVar.Filters select new
                        {
                            name = f.Name,
                            args = f.Arguments
                        });

                        if (tonkVar.Filters.Count > 0)
                            currentNode.Add(new JProperty("filters", filters));
                        break;
                }
            }
        }
    }
}
