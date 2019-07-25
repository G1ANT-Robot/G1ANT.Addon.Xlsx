using System;
using System.Linq;

namespace G1ANT.Addon.Xlsx.Api
{
    public class CellRef : IEquatable<CellRef>
    {
        public CellRef(string sheetId, string col, int row)
        {
            SheetId = sheetId;
            Column = col.ToUpper();
            Row = row;
        }

        public CellRef(string sheetId, string address)
        {
            SheetId = sheetId;

            var firstDigitPosition = address.IndexOfAny("0123456789".ToCharArray());

            Column = address.Substring(0, firstDigitPosition);
            Row = int.Parse(address.Substring(firstDigitPosition));
        }

        public string SheetId { get; }
        public string Column { get; }
        public int Row { get; }

        public string Address => $"{Column}{Row}";

        public CellRef[] BuildMatrix(CellRef otherCorner)
        {
            if (SheetId != otherCorner.SheetId)
            {
                return null;
            }

            if (this == otherCorner)
            {
                return new CellRef[]
                {
                    new CellRef(this.SheetId, this.Column, this.Row)
                };
            }

            var startColumn = Column.CompareTo(otherCorner.Column) < 0 ? Column : otherCorner.Column;
            var endColumn = Column.CompareTo(otherCorner.Column) > 0 ? Column : otherCorner.Column;
            var startRow = Row < otherCorner.Row ? Row : otherCorner.Row;
            var endRow = Row > otherCorner.Row ? Row : otherCorner.Row;

            var start = new CellRef(SheetId, startColumn, startRow);
            var end = new CellRef(SheetId, endColumn, endRow);

            long rows = Math.Abs(start.Row - end.Row) + 1;
            var columns = new ColumnNamesGenerator().GetColumnsBetweenInclusive(start.Column, end.Column);

            var result = new CellRef[rows * columns.Length];

            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new CellRef(SheetId, columns[i % columns.Length], i / columns.Length + start.Row);
            }

            return result;
        }

        public bool Equals(CellRef other)
        {
            if (ReferenceEquals(null, other))
            {
                return false;
            }

            if (ReferenceEquals(this, other))
            {
                return true;
            }

            return Address == other.Address && SheetId == other.SheetId;
        }

        public static bool operator ==(CellRef cell1, CellRef cell2)
        {
            return cell1.Equals(cell2);
        }

        public static bool operator !=(CellRef cell1, CellRef cell2)
        {
            return !cell1.Equals(cell2);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as CellRef);
        }

        public override int GetHashCode()
        {
            return Address.GetHashCode() ^ SheetId.GetHashCode();
        }

        public override string ToString()
        {
            return Address;
        }
    }
}
