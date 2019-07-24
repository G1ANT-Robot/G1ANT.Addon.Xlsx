using System;

namespace G1ANT.Addon.Xlsx.Api
{
    public class CellR
    {
        public CellR(string col, int row)
        {
            Column = col.ToUpper();
            Row = row;
        }

        public string Column { get; }
        public int Row { get; }

        public string Address => $"{Column}{Row}";

        public CellR[] BuildMatrix(CellR otherCorner)
        {
            if (this == otherCorner)
            {
                return new CellR[1]
                {
                    new CellR(this.Column, this.Row)
                };
            }

            var startColumn = this.Column.CompareTo(otherCorner.Column) < 0 ? this.Column : otherCorner.Column;
            var endColumn = this.Column.CompareTo(otherCorner.Column) > 0 ? this.Column : otherCorner.Column;
            var startRow = this.Row < otherCorner.Row ? this.Row : otherCorner.Row;
            var endRow = this.Row > otherCorner.Row ? this.Row : otherCorner.Row;

            var start = new CellR(startColumn, startRow);
            var end = new CellR(endColumn, endRow);

            long rows = Math.Abs(start.Row - end.Row) + 1;
            var columns = new ColumnNamesGenerator().GetColumnsBetweenInclusive(start.Column, end.Column);

            var result = new CellR[rows * columns.Length];

            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new CellR(columns[i % columns.Length], i / columns.Length + start.Row);
            }

            return result;
        }

        public override string ToString()
        {
            return Address;
        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return Address.GetHashCode();
        }
    }
}
