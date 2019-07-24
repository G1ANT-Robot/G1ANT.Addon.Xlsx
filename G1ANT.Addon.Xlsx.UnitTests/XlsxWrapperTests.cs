using G1ANT.Addon.Xlsx.Api;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxWrapperTests
    {
        [Test]
        public void TestSelection()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(new CellR("B", 3), new CellR("D", 6));
            Assert.AreEqual(wrapper.SelectedCells.Count(), 12);
        }

        [Test]
        public void TestSelectionWrongPairs()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(new CellR("C", 4), new CellR("A", 8));
            Assert.AreEqual(wrapper.SelectedCells.Count(), 15);
        }
    }
}
