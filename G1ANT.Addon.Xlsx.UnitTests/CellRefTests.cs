using NUnit.Framework;
using G1ANT.Addon.Xlsx.Api;

namespace G1ANT.Addon.Xlsx.UnitTests
{
    [TestFixture]
    public class CellRefTests
    {
        [Test]
        public void TestConstructor()
        {
            CellRef a = new CellRef("sheet", "ABC42");

            Assert.AreEqual("ABC42", a.Address);
            Assert.AreEqual(42, a.Row);
            Assert.AreEqual("ABC", a.Column);
        }

        [Test]
        public void TestCellRefsEquality()
        {
            CellRef a = new CellRef("sheet", "A", 4);
            CellRef b = new CellRef("sheet", "A", 4);
            object c = new CellRef("sheet", "A", 4);

            Assert.AreEqual(a, b);
            Assert.IsTrue(a == b);
            Assert.IsTrue(a.Equals(b));
            Assert.IsTrue(a.Equals((object)b));

            Assert.IsFalse(c == a);
            Assert.IsTrue(c.Equals(a));
        }
    }
}
