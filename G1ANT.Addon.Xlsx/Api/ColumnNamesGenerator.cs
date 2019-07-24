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
using System.Text;

namespace G1ANT.Addon.Xlsx.Api
{
    public class ColumnNamesGenerator
    {
        private readonly char[] letters;
        private readonly string[] columnNames;

        public ColumnNamesGenerator()
        {
            letters = Enumerable.Range(0, 26).Select(x => (char)(x + 64)).ToArray();
            columnNames = Enumerable
                .Range(0, ushort.MaxValue)
                .Select(x => CalcString(x))
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

            var result = columnNames
                .SkipWhile(x => x != start)
                .TakeWhile(x => x != end)
                .ToList();

            result.Add(end);

            return result.ToArray();
        }

        private string CalcString(int index)
        {
            StringBuilder sb = new StringBuilder();

            while (index > 0)
            {
                sb.Append(letters[index % letters.Length]);
                index /= letters.Length;
            }

            return new string(sb.ToString().Reverse().ToArray());
        }
    }
}

