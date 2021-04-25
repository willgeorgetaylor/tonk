using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Xml;
using Newtonsoft.Json.Linq;
using Syncfusion.DocIO;

namespace Tonk.DOM
{
    internal static class Rehydrator
    {
        private static void Cleanup(XmlDocument doc)
        {
            var matchingElements = doc.SelectNodes(@"//WParagraph[Text]");
            if (matchingElements == null) return;

            var toRemove = new List<string>();
            foreach (XmlNode par in matchingElements)
            {
                bool hasTonkSnippet = false;
                bool allEmpty = true;

                foreach (XmlNode childNode in par.ChildNodes)
                {
                    if (childNode.Attributes["tonk"] != null)
                    {
                        hasTonkSnippet = true;
                    }

                    if (childNode.Name == "Text" && childNode.HasChildNodes && childNode.FirstChild.NodeType == XmlNodeType.Text && childNode.FirstChild.InnerText.Length > 0 )
                    {
                        allEmpty = false;
                        break;
                    }
                }

                if (hasTonkSnippet && allEmpty)
                {
                    par.ParentNode.RemoveChild(par);
                    //if (par.Attributes != null) toRemove.Add(par.Attributes["id"].ToString());
                }

            }
        }

        private static void RemoveTableHelpers(XmlDocument doc)
        {
            var tonkHelperLeftovers = doc.SelectNodes(@"//WTable/Text[@tonk='yes']");
            if (tonkHelperLeftovers == null) return;

            foreach (XmlNode tonk in tonkHelperLeftovers)
            {
                tonk.ParentNode.RemoveChild(tonk);
            }
        }

        public static WordDocument Rehydrate(string newDOM, WordDocument templateDoc, Dictionary<int, IEntity> elements)
        {
            var doc = new XmlDocument {PreserveWhitespace = true};
            doc.LoadXml(newDOM);
            RemoveTableHelpers(doc);

            var blankDoc = templateDoc.Clone();
            blankDoc.ChildEntities.Clear();

            var x = blankDoc.AddSection();
            blankDoc.ChildEntities.RemoveAt(0);


            foreach (var traverseNode in TraverseNodes(doc.ChildNodes[0].ChildNodes, blankDoc, blankDoc, elements))
            {
          
            }

            return blankDoc;
        }

        private static IEnumerable<IEntity> TraverseNodes(XmlNodeList nodes, WordDocument doc, ICompositeEntity seedEntity, IReadOnlyDictionary<int, IEntity> elements)
        {
            for (var index = 0; index < nodes.Count; index++)
            {
                XmlNode node = nodes[index];
                IEntity entity;

                // Either we have a stored entity we can just restore
                if (node.Attributes != null && node.Attributes.Count > 0)
                {
                    var id = Convert.ToInt32(node.Attributes["id"].Value);
                    IEntity item;
                    if (elements.TryGetValue(id, out item))
                    {
                        Console.WriteLine("Found match for entity " + id + " (" + elements[id].EntityType.ToString() +
                                          ")");
                        entity = elements[id].Clone();
                    }
                    else
                    {
                        entity = new WTextRange(doc);
                    }
                }
                else
                {
                    // Or it's a new text range (rare)
                    entity = new WTextRange(doc);
                }

                // Not a placeholder
                if (node.HasChildNodes)
                {
                    if (node.Name == "Text" &&
                        (node.FirstChild.NodeType == XmlNodeType.Text ||
                         node.FirstChild.NodeType == XmlNodeType.Whitespace ||
                         node.FirstChild.NodeType == XmlNodeType.SignificantWhitespace) &&
                        entity.GetType() == typeof(WTextRange))
                    {
                        if (node.Attributes["tonk"] != null &&
                            !(node.HasChildNodes && node.FirstChild.NodeType == XmlNodeType.Text &&
                              node.FirstChild.InnerText.Length > 0))
                        {
                        }
                        else
                        {
                            var textRange = entity as WTextRange;

                            var imageTagResult = ImageTagFinder.Match(node.FirstChild.InnerText);
                            var linkTagResult = LinkTagFinder.Match(node.FirstChild.InnerText);

                            if (imageTagResult.Success)
                            {
                                var parentPar = seedEntity as WParagraph;
                                try
                                {
                                    string localFilename =
                                        @"C:\Users\Will\Desktop\image-to-insert-" + DateTime.Now.Ticks.ToString() +
                                        ".jpg";

                                    using (var client = new WebClient())
                                    {
                                        client.DownloadFile(imageTagResult.Groups[1].Value, localFilename);
                                    }

                                    var picture = parentPar.AppendPicture(Image.FromFile(localFilename));
                                    picture.Height = 100;
                                    picture.Width = 100;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }
                            else if (linkTagResult.Success)
                            {
                                var parentPar = seedEntity as WParagraph;
                                
                               
                                IWCharacterStyle myStyle = doc.AddCharacterStyle("sdfds");


                                //Sets the formatting of the style

                                myStyle.CharacterFormat.FontSize = 16f;

                                myStyle.CharacterFormat.TextBackgroundColor = Color.Red;
                                
                                
                                WField field = parentPar.AppendField("Hyperlink", FieldType.FieldHyperlink) as WField;
                                Hyperlink link = new Hyperlink(field);
                                link.Type = HyperlinkType.WebLink;
                                link.Uri = linkTagResult.Groups[1].Value;

                            }
                            else
                            {
                                textRange.Text = node.FirstChild.InnerText;
                                seedEntity.ChildEntities.Add(textRange);
                                yield return textRange;
                            }
                        }
                    }
                    else
                    {
                        var parentEntity = entity as ICompositeEntity;


                        if (seedEntity.GetType() == typeof(WSection))
                        {
                            var seedSection = seedEntity as WSection;

                            if (parentEntity.GetType() == typeof(HeaderFooter))
                            {
                                foreach (var childEntity in TraverseNodes(node.ChildNodes, doc, seedSection.HeadersFooters[index - 1], elements))
                                {
                                    yield return childEntity;
                                }
                            }
                            else
                            {
                                foreach (var childEntity in TraverseNodes(node.ChildNodes, doc, seedSection.Body,
                                    elements))
                                {
                                    yield return childEntity;
                                }
                            }
                        }
                        else
                        {
                            seedEntity.ChildEntities.Add(parentEntity);

                            // This is a parent node
                            foreach (var childEntity in TraverseNodes(node.ChildNodes, doc, parentEntity, elements))
                            {
                                yield return childEntity;
                            }
                        }
                    }
                }
                else
                {
                    if (seedEntity.GetType() != typeof(WSection))
                    {
                        // This is a placeholder node
                        seedEntity.ChildEntities.Add(entity);
                        yield return entity;
                    }
                }
            }
        }

        private static readonly Regex ImageTagFinder = new Regex(@"^image\{(.+)\}$");
        private static readonly Regex LinkTagFinder = new Regex(@"^link\{(.+)\}$");
    }
}