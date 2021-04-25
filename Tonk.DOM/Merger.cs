using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DotLiquid;
using Newtonsoft.Json.Linq;

namespace Tonk.DOM
{
    internal static class Merger
    {
        private static readonly Regex ExtractCommandsRegex = new Regex(@"<Text.*?>(\{%.+?%\})<\/Text>");


        public static string Merge(string json, string dom)
        {
            Template.RegisterFilter(typeof(ImageFilter));
            Template.RegisterFilter(typeof(LinkFilter));

            var template = Template.Parse(dom);  // Compile the template
            var drop = Dropifier.Dropify(json);  // Convert JSON => DotLiquid Hash
            return template.Render(drop);        // Fill fields
        }
    }


    public static class ImageFilter
    {
        public static string Imagify(object input, string format, string wazoo)
        {
            if (input == null)
                return null;
            else if (string.IsNullOrWhiteSpace(format))
                return input.ToString();

            Console.WriteLine(format);
            Console.WriteLine(wazoo);
            return "image{" + input.ToString() + "}";

        }
    }

    public static class LinkFilter
    {
        public static string Linkify(Context context, string input)
        {
            Console.WriteLine(context);
            Console.WriteLine(context["height"]);
            return "link{" + input + "}";
        }
    }
}


