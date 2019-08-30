/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System.Linq;

namespace G1ANT.Addon.Xlsx.Api
{
    public class ColumnNamesGenerator
    {
        private static char[] Letters;
        private static string[] ColumnNames;

        static ColumnNamesGenerator()
        {
            Letters = Enumerable.Range(0, 26).Select(x => (char)(x + 64)).ToArray();
            ColumnNames = Enumerable
                .Range(0, ushort.MaxValue)
                .Select(x => GenerateColumn(x))
                .ToArray();
        }
           
        public string[] GetColumnsBetweenInclusive(string column1, string column2)
        {
            var start = column1;
            var end = column2;

            if (start.CompareTo(end) > 0)
            {
                var b = end;
                end = start;
                start = b;
            }

            var result = ColumnNames
                .SkipWhile(x => x != start)
                .TakeWhile(x => x != end)
                .ToList();

            result.Add(end);

            return result.ToArray();
        }

        private static string GenerateColumn(int index)
        {
            string column = "";

            while (index > 0)
            {
                column += Letters[index % Letters.Length];
                index /= Letters.Length;
            }

            return new string(column.Reverse().ToArray());
        }
    }
}

