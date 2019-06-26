using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Addon.Xlsx.Tests.Properties;
using G1ANT.Engine;
using G1ANT.Language;
using NUnit.Framework;



namespace G1ANT.Addon.Xlsx.Tests.Commands
{
    [TestFixture]
    public class GetValuesInRangeTests
    {
        Scripter scripter;
        string file;
        [OneTimeSetUp]
        [Timeout(20000)]
        public void ClassInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            scripter = new Scripter();
            scripter.InitVariables.Clear();
        }

        [Test]
        public void GetValueInRangeTest()
        {

        }
    }
}
