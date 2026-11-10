using System;
using System.Collections.Generic;

namespace NPOI.OpenXmlFormats.Presentation
{
    internal static class PresentationXmlContent
    {
        internal static object[] Merge(int[] order, params object[][] groups)
        {
            var result = new List<object>();
            var positions = new int[groups.Length];
            var lastPositions = new int[groups.Length];
            for (var group = 0; group < groups.Length; group++)
            {
                lastPositions[group] = Array.LastIndexOf(order, group);
            }

            void AppendRemaining(int group)
            {
                var values = groups[group] ?? Array.Empty<object>();
                while (positions[group] < values.Length)
                {
                    var value = values[positions[group]++];
                    if (value != null)
                    {
                        result.Add(value);
                    }
                }
            }

            for (var index = 0; index < order.Length; index++)
            {
                var group = order[index];
                // Insert newly modeled children in schema order without moving opaque XML.
                if (group < groups.Length - 1)
                {
                    for (var preceding = 0; preceding < group; preceding++)
                    {
                        if (lastPositions[preceding] < 0)
                        {
                            AppendRemaining(preceding);
                        }
                    }
                }
                var values = groups[group] ?? Array.Empty<object>();
                if (positions[group] < values.Length)
                {
                    var value = values[positions[group]++];
                    if (value != null)
                    {
                        result.Add(value);
                    }
                }
                if (lastPositions[group] == index)
                {
                    AppendRemaining(group);
                }
            }
            for (var group = 0; group < groups.Length; group++)
            {
                AppendRemaining(group);
            }
            return result.ToArray();
        }
    }
}
