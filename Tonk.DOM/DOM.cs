using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Syncfusion.DocIO.DLS;
using DotLiquid;
using Newtonsoft.Json.Linq;
using Template = DotLiquid.Template;

namespace Tonk.DOM
{
    public static class DOM
    {
        public static string Parse(WordDocument templateDoc)
        {
            var cleanDoc = Prepper.Prep(templateDoc);
            string templateDOM = Dehydrator.Dehydrate(cleanDoc);

            Template.RegisterFilter(typeof(ImageFilter));
            Template.RegisterFilter(typeof(LinkFilter));

            var parse = Template.Parse(templateDOM);
            var cc = Parser.Parse(parse.Root);
            return cc;
        }

        public static WordDocument Merge(WordDocument templateDoc, string json)
        {
            var cleanDoc = Prepper.Prep(templateDoc);
            var entityDict = Extractor.ExtractEntities(cleanDoc);
            string templateDOM = Dehydrator.Dehydrate(cleanDoc);         // Convert Doc => Tonk
            Console.WriteLine("======== OLD DOM =========");
            Console.WriteLine(Dehydrator.PrintXml(templateDOM));
            string newDOM = Merger.Merge(json, templateDOM);                // Do the merge on the Tonk
            Console.WriteLine("======== NEW DOM =========");
            Console.WriteLine(Dehydrator.PrintXml(newDOM));
            var newDoc = Rehydrator.Rehydrate(newDOM, cleanDoc, entityDict);   // Convert Tonk => Doc
            return newDoc;
        }

    }
}