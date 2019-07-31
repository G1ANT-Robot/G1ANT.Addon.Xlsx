using G1ANT.Addon.Xlsx.Api;
using NUnit.Framework;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace G1ANT.Addon.Xlsx.UnitTests
{
    [TestFixture]
    public class XlsxWrapperTests
    {
        [Test]
        public void TestSelection()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(new CellRef("sheet", "B", 3), new CellRef("sheet", "D", 6));
            Assert.AreEqual(12, wrapper.SelectedCells.Count());
        }

        [Test]
        public void TestSelectionOnWrongPairs()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(new CellRef("sheet", "C", 4), new CellRef("sheet", "A", 8));
            Assert.AreEqual(15, wrapper.SelectedCells.Count());
        }

        [Test]
        public void TestSelectionOnDifferentSheets()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(new CellRef("sheet1", "C", 4), new CellRef("sheet2", "A", 8));
            Assert.AreEqual(null, wrapper.SelectedCells);
        }

        [Test]
        public void TestGetColor()
        {
            var wrapper = new XlsxWrapper(0);

            //wrapper.GetCellColor(new CellRef("sheet1", "C", 4));
            Assert.AreEqual(null, wrapper.SelectedCells);
        }
    }
}
