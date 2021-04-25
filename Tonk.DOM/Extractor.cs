using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Drawing;

namespace Tonk.DOM
{
    internal static class Extractor
    {
        public static Dictionary<int, IEntity> ExtractEntities(WordDocument templateDoc)
        {
            var elements = new Dictionary<int, IEntity>();
            foreach (var obj in TraverseParentEntity(templateDoc))
            {
                var old = obj.GetHashCode().ToString();
                var copyMe = obj.Clone();
                var nn = copyMe.GetHashCode().ToString();

                if (copyMe is ICompositeEntity possibleParentEntity)
                {
                    if (possibleParentEntity.GetType() == typeof (WSection))
                    {
                        var parentSec = possibleParentEntity as WSection;
                        parentSec.Body.ChildEntities.Clear();
                        foreach (HeaderFooter hf in parentSec.HeadersFooters)
                        {
                            hf.ChildEntities.Clear();
                        }
                    }
                    else
                    {
                        possibleParentEntity.ChildEntities.Clear();
                    }
                    
                    elements[obj.GetHashCode()] = possibleParentEntity;
                }
                else
                {
                    if (copyMe.GetType() == typeof (WTextRange))
                    {
                        var original = obj as WTextRange;
                        var tr = copyMe as WTextRange;
                        
                        var characterformat = new WCharacterFormat(templateDoc)
                        {
                            Bold = original.CharacterFormat.Bold,
                            Italic = original.CharacterFormat.Italic,
                            TextColor = original.CharacterFormat.TextColor,
                            Font = original.CharacterFormat.Font
                        };

                        tr.ApplyCharacterFormat(characterformat);
                        tr.Text = "";
                        elements[obj.GetHashCode()] = tr;
                    }
                    else
                    {
                        elements[obj.GetHashCode()] = obj;
                    }
                }
            }
            return elements;
        }

        public static IEnumerable<IEntity> TraverseParentEntity(ICompositeEntity parentEntity)
        {
            if (parentEntity == null) yield break;

            foreach (var childEntity in parentEntity.ChildEntities)
            {
                yield return childEntity as IEntity;

                var possibleParentEntity = childEntity as ICompositeEntity;
                if (possibleParentEntity == null) continue;
                foreach (var grandchildEntity in TraverseParentEntity(possibleParentEntity))
                {
                    yield return grandchildEntity;
                }
            }
        }
    }
}