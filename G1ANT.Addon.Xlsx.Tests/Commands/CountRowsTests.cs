/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Addon.Xlsx.Tests.Properties;
using G1ANT.Engine;
using G1ANT.Language;
using NUnit.Framework;
using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class CountRowsTests
    {
        Scripter scripter;
        string file;
        string file2;

        [OneTimeSetUp]
        [Timeout(20000)]
        public void Initialize()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            file2 = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.EmptyWorkbook), "xlsx");
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            scripter = new Scripter();
        }
       
        [Test]
        [Timeout(20000)]
        public void CountRowsTest()
        {
            int rowCount;
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(file));
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.countrows result {SpecialChars.Variable}rowCunt";
            scripter.Run();
            rowCount = scripter.Variables.GetVariableValue<int>("rowCunt", -1, true);
            Assert.AreEqual(5, rowCount);
        }

        [Test]
        [Timeout(20000)]
        public void CountRowsInEmptyWoorkbookTest()
        {
            int rowCount;
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(file2));
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.countrows result {SpecialChars.Variable}rowCount";
            scripter.Run();
            rowCount = scripter.Variables.GetVariableValue<int>("rowCount", -1, true);
            Assert.AreEqual(0, rowCount);
        }

        [OneTimeTearDown]
        [Timeout(10000)]
        public void ClassCleanUp()
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
            if (File.Exists(file2))
            {
                File.Delete(file2);
            }
        }
    }
}
