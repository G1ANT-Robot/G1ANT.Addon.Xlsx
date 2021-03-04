using G1ANT.Addon.Xlsx.Api;
using NUnit.Framework;
using System.Linq;

namespace G1ANT.Addon.Xlsx.UnitTests
{
    [TestFixture]
    public class XlsxWrapperTests
    {
        [Test]
        public void ShouldSelectedCellsUpdate_WhenCaliingSelectRange()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(3, "B", 6, "D");
            Assert.AreEqual(12, wrapper.SelectedCells.Cells().Count());
        }

        [Test]
        public void ShouldSelectionWork_WhenMixingOrderOfSelectionCorners()
        {
            var wrapper = new XlsxWrapper(0);

            wrapper.SelectRange(4, "C", 8 , "A");
            Assert.AreEqual(15, wrapper.SelectedCells.Cells().Count());
        }
    }
}
