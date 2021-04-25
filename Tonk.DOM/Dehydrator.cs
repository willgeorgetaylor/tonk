using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using HeaderFooter = Syncfusion.DocIO.DLS.HeaderFooter;

namespace Tonk.DOM
{
    public static class Dehydrator
    {
        public static String PrintXml(String xml)
        {
            String result = "";

            var mStream = new MemoryStream();
            var writer = new XmlTextWriter(mStream, Encoding.Unicode);
            var document = new XmlDocument();
            
            try
            {
                // Load the XmlDocument with the XML.
                document.LoadXml(xml);

                writer.Formatting = Formatting.Indented;

                // Write the XML into a formatting XmlTextWriter
                document.WriteContentTo(writer);
                writer.Flush();
                mStream.Flush();

                // Have to rewind the MemoryStream in order to read
                // its contents.
                mStream.Position = 0;

                // Read MemoryStream contents into a StreamReader.
                var sReader = new StreamReader(mStream);

                // Extract the text from the StreamReader.
                string formattedXml = sReader.ReadToEnd();

                result = formattedXml;
            }
            catch (XmlException)
            {
            }

            mStream.Close();
            writer.Close();

            return result;
        }

        private static XmlElement Placeholder(XmlDocument doc, object obj)
        {
            var placeholderElement = doc.CreateElement(obj.GetType().Name);

            if (obj.GetType() == typeof(WParagraph))
            {
                var par = obj as WParagraph;
                //Assign the ID attribute
                var sattr = doc.CreateAttribute("style");
                sattr.Value = par.StyleName;
                placeholderElement.SetAttributeNode(sattr);
            }

            //Assign the ID attribute
            var attr = doc.CreateAttribute("id");
            attr.Value = obj.GetHashCode().ToString();

            placeholderElement.SetAttributeNode(attr);

            return placeholderElement;
        }

        private static XmlElement Text(XmlDocument doc, WTextRange text, bool tonkCommand = false)
        {
            var textElement = doc.CreateElement("Text");

            //Assign the ID attribute
            var attr = doc.CreateAttribute("id");
            attr.Value = text.GetHashCode().ToString();
            textElement.SetAttributeNode(attr);

            //Assign the style attribute
            var sattr = doc.CreateAttribute("style");
            sattr.Value = text.CharacterFormat.CharStyleName?.ToString();
            textElement.SetAttributeNode(sattr);

            //Assign the style attribute
            var fattr = doc.CreateAttribute("font");
            fattr.Value = text.CharacterFormat.FontName.ToString();
            textElement.SetAttributeNode(fattr);

            //Assign the style attribute
            var xattr = doc.CreateAttribute("size");
            xattr.Value = text.CharacterFormat.FontSize.ToString(CultureInfo.InvariantCulture);
            textElement.SetAttributeNode(xattr);


            if (tonkCommand)
            {
                //In case it's a Tonk snippet
                var attr2 = doc.CreateAttribute("tonk");
                attr2.Value = "yes";
                textElement.SetAttributeNode(attr2);
            }
 
            textElement.InnerText = text.Text;

            return textElement;
        }

        private static IEnumerable<IEntity> DehydrateRecursive(ICompositeEntity parentEntity, XmlNode parentNode, XmlDocument xmldoc)
        {
            foreach (var childEntity in parentEntity.ChildEntities)
            {
                var castEntity = childEntity as IEntity;
                yield return castEntity;

                var childNode = Render(castEntity, xmldoc);
                parentNode.AppendChild(childNode);

                var possibleParentEntity = childEntity as ICompositeEntity;

                if (possibleParentEntity == null) continue; // It doesn't have child entities
                if (possibleParentEntity.GetType() == typeof(WParagraph)) continue; // It's a paragraph node, render will take over from here

                foreach (var grandchildEntity in DehydrateRecursive(possibleParentEntity, childNode, xmldoc))
                {
                    yield return grandchildEntity;
                }
            }
        }

        private static XmlElement Render(IEntity entity, XmlDocument xmldoc)
        {
            //if (entity.GetType() == typeof(WSection))
            //{
            //    var section = entity as WSection;
            //    var sectionElement = Placeholder(xmldoc, section);

            //    for (int j = 0; j < 6; j++)
            //    {
            //        var headerElement = Text(xmldoc, tRange, isTonkSnippet);
            //        paragraphElement.AppendChild(textElement);
            //    }

            //    return sectionElement;
            //}

            if (entity.GetType() != typeof(WParagraph))
            {
                return Placeholder(xmldoc, entity);
            }

            var paragraph = (WParagraph)entity;
            var paragraphElement = Placeholder(xmldoc, paragraph);

            foreach (var parChild in paragraph.ChildEntities)
            {
                if (parChild.GetType() == typeof(WTextRange))
                {
                    var tRange = parChild as WTextRange;
                    bool isTonkSnippet = tRange.Text.StartsWith("{%") && tRange.Text.EndsWith("%}");
                    var textElement = Text(xmldoc, tRange, isTonkSnippet);
                    paragraphElement.AppendChild(textElement);
                }
                else
                {
                    paragraphElement.AppendChild(Placeholder(xmldoc, parChild));
                }
            }

            return paragraphElement;
        }

        private static void FormatTables(XmlDocument doc)
        {
            var tonkStartsInCells = doc.SelectNodes(@"//WTable/WTableRow/WTableCell[1]/WParagraph/Text[@tonk='yes']");
            var tonkEndsInCells = doc.SelectNodes(@"//WTable/WTableRow/WTableCell[last()]/WParagraph/Text[@tonk='yes']");

            if (tonkStartsInCells == null || tonkEndsInCells == null) return;

            foreach (XmlNode starter in tonkStartsInCells)
            {
                var parentTable = starter.ParentNode.ParentNode.ParentNode.ParentNode;
                var parentRow = starter.ParentNode.ParentNode.ParentNode;
                parentTable?.InsertBefore(starter, parentRow);
            }

            foreach (XmlNode ender in tonkEndsInCells)
            {
                var parentTable = ender.ParentNode.ParentNode.ParentNode.ParentNode;
                var parentRow = ender.ParentNode.ParentNode.ParentNode;
                parentTable?.InsertAfter(ender, parentRow);
            }
        }

        public static string Dehydrate(WordDocument templateDoc)
        {
            var doc = new XmlDocument();
            var documentRoot = doc.CreateElement("doc");
            doc.AppendChild(documentRoot);
            
            foreach (var c in DehydrateRecursive(templateDoc, documentRoot, doc))
            {
            }
            FormatTables(doc);
            return doc.OuterXml;
        }
    }
}