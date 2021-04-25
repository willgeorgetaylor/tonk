using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Tonk.DOM
{
    internal static class Prepper
    {
        public static WordDocument Prep(WordDocument rawDoc)
        {
            var cleanDoc = rawDoc.Clone();
            PrepRecursive(cleanDoc);
            return cleanDoc;
        }

        private static readonly Type[] InvisibleRangeBlockers = { typeof(Bookmark), typeof(BookmarkStart), typeof(BookmarkEnd) };

        private static void PrepRecursive(ICompositeEntity parentEntity)
        {
            foreach (var childEntity in parentEntity.ChildEntities)
            {
                if (childEntity.GetType() == typeof(WParagraph))
                {
                    var paragraph = (WParagraph)childEntity;
                    var newParContents = new List<IEntity>();
                    

                    var savedUp = new List<WTextRange>();

                    foreach (IEntity parChild in paragraph.ChildEntities)
                    {
                        if (InvisibleRangeBlockers.Contains(parChild.GetType()))
                        {
                            continue;
                        }

                        if (parChild.GetType() == typeof(WTextRange))
                        {
                            var textRange = (WTextRange)parChild;
                            savedUp.Add(textRange);
                        }
                        else
                        {
                            // OK, random element, Imma let you finish but first...
                            if (savedUp.Count > 0)
                            {
                                foreach (var textRange in SplitRanges(savedUp))
                                {
                                    newParContents.Add(textRange);
                                }
                            }

                            newParContents.Add(parChild);
                        }
                    }

                    if (savedUp.Count > 0)
                    {
                        foreach (var textRange in SplitRanges(savedUp))
                        {
                            newParContents.Add(textRange);
                        }
                    }
                    
                    paragraph.ChildEntities.Clear();
                    foreach (var ent in newParContents)
                    {
                        paragraph.ChildEntities.Add(ent);
                    }
                }
                else
                {
                    var possibleParentEntity = childEntity as ICompositeEntity;
                    if (possibleParentEntity == null) continue;
                    PrepRecursive(possibleParentEntity);
                }
            }
        }

        private static readonly Regex LiquidStartFinder = new Regex("(?<!{){");
        private static readonly Regex LiquidEndFinder = new Regex("}(?!})");

        private static IEnumerable<WTextRange> SplitRanges(IReadOnlyCollection<WTextRange> input)
        {
            Console.WriteLine("======= Split PAR");

            if (input.Count == 0)
            {
                return null;
            }

            var newRanges = new List<WTextRange>();

            foreach (var textRange in input)
            {
                string rangeContents = textRange.Text;

                var startPoints = (from Match match in LiquidStartFinder.Matches(rangeContents) select match.Index).ToList();
                var endPoints = (from Match match in LiquidEndFinder.Matches(rangeContents) select match.Index + 1).ToList();
                var splitPoints = startPoints.Union(endPoints).OrderBy(num => num).ToList();

                if (splitPoints.Count > 0)
                {
                    int lastIndex = 0;
                    foreach (int splitter in splitPoints)
                    {
                        var newR = (WTextRange)textRange.Clone();
                        newR.Text = rangeContents.Substring(lastIndex, splitter - lastIndex);
                        newRanges.Add(newR);
                        //Save our position
                        lastIndex = splitter;
                    }

                    if (lastIndex >= rangeContents.Length) continue;
                    var lastR = (WTextRange)textRange.Clone();
                    lastR.Text = rangeContents.Substring(lastIndex, (rangeContents.Length - lastIndex));
                    newRanges.Add(lastR);
                }
                else
                {
                    // No Liquid in here, pass through
                    newRanges.Add(textRange);
                }
            }

            Console.WriteLine("======= POST Split");
            foreach (var range in newRanges)
            {
                Console.WriteLine("Range is '" + range.Text + "' (" + range.CharacterFormat.FontName.ToString() + ")");
            }

            return ConsolidateRanges(newRanges);
        }

        private static bool TerminalBrace(IList<WTextRange> ranges, int braceRangeIndex)
        {
            if (braceRangeIndex == ranges.Count - 1) return true;
            for (int idx = braceRangeIndex + 1; idx < ranges.Count; idx++)
            {
                var rangeText = ranges[idx].Text;
                if (rangeText.Length == 0) continue;
                return (!rangeText.StartsWith("}"));
            }
            return true;
        }

        private static IEnumerable<WTextRange> ConsolidateRanges(IList<WTextRange> ranges)
        {
            if (ranges.Count == 0)
            {
                return null;
            }

            var newRanges = new List<WTextRange>();
            WTextRange chain = null;

            for (int idx = 0; idx < ranges.Count; idx++)
            {
                // Only start a new chain when we've terminated successfully
                if (ranges[idx].Text.StartsWith("{"))
                {
                    if (ranges[idx].Text.EndsWith("}"))
                    {
                        newRanges.Add(ranges[idx]);
                    }
                    else
                    {
                        if (chain == null)
                        {
                            chain = ranges[idx];
                        }
                        else
                        {
                            chain.Text = chain.Text + ranges[idx].Text;
                        }
                    }
                
                }
                // When we reach end braces
                else if (ranges[idx].Text.EndsWith("}"))
                {
                    // Very end brace? Continue until can prove otherwise, or we hit the end of ranges in this paragraph.
                    if (TerminalBrace(ranges, idx))
                    {
                        if (chain == null) continue; // Nothing to terminate, so that's weird
                        chain.Text = chain.Text + ranges[idx].Text;
                        newRanges.Add(chain);
                        chain = null;
                    }
                    else
                    {
                        ChainOrPassthrough(chain, newRanges, ranges[idx]);
                    }
                }
                else
                {
                    ChainOrPassthrough(chain, newRanges, ranges[idx]);
                }
            }

            if (chain != null)
            {
                newRanges.Add(chain);
            }

            Console.WriteLine("======= POST Consolidation");
            foreach (var range in newRanges)
            {
                Console.WriteLine("Range is '" + range.Text + "' (" + range.CharacterFormat.FontName.ToString() + ")");
            }

            return newRanges;
        }

        private static void ChainOrPassthrough(WTextRange chain, ICollection<WTextRange> newRanges, WTextRange currentRange)
        {
            if (chain != null)
            {
                chain.Text = chain.Text + currentRange.Text;
            }
            else
            {
                newRanges.Add(currentRange);
            }
        }
    }
}