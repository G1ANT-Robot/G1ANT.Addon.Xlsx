/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;
using System.Collections.Generic;
using System.IO;

using G1ANT.Engine;

using NUnit.Framework;
using System.Reflection;
using G1ANT.Language;
using G1ANT.Addon.Xlsx.Tests.Properties;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxFindTest
    {
        string file;
        string file2;
        Scripter scripter;
        [OneTimeSetUp]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            file2 = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.EmptyWorkbook), "xlsx");
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(file));
        }

        [Test]
        [Timeout(40000)]
        public void XlsxFindDifferentTypesTest()
        {

            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.find 1234 resultrow {SpecialChars.Variable}resrow resultcolumn {SpecialChars.Variable}resCol
            xlsx.find {SpecialChars.Text}abcd{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow2 resultcolumn {SpecialChars.Variable}resCol2
            xlsx.find 150 resultrow {SpecialChars.Variable}resrow3 resultcolumn {SpecialChars.Variable}resCol3
            -xlsx.find {SpecialChars.Text}160%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow4 resultcolumn {SpecialChars.Variable}resCol4
            -xlsx.find {SpecialChars.Text}100%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow5 resultcolumn {SpecialChars.Variable}resCol5
            xlsx.find {SpecialChars.Text}AA{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow6 resultcolumn {SpecialChars.Variable}resCol6
            xlsx.find {SpecialChars.Text}AZ{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow7 resultcolumn {SpecialChars.Variable}resCol7
            xlsx.find {SpecialChars.Text}BA{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow8 resultcolumn {SpecialChars.Variable}resCol8
            xlsx.find {SpecialChars.Text}AAZ{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow9 resultcolumn {SpecialChars.Variable}resCol9
            xlsx.find {SpecialChars.Text}ABC{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow10 resultcolumn {SpecialChars.Variable}resCol10
            xlsx.find {SpecialChars.Text}Z{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow11 resultcolumn {SpecialChars.Variable}resCol11

";
            scripter.Run();
            Assert.AreEqual(1, scripter.Variables.GetVariable("resrow").GetValue().Object);
            Assert.AreEqual(1, scripter.Variables.GetVariable("resCol").GetValue().Object);

            Assert.AreEqual(1, scripter.Variables.GetVariable("resrow2").GetValue().Object);
            Assert.AreEqual(2, scripter.Variables.GetVariable("resCol2").GetValue().Object);

            Assert.AreEqual(1, scripter.Variables.GetVariable("resrow3").GetValue().Object);
            Assert.AreEqual(4, scripter.Variables.GetVariable("resCol3").GetValue().Object);

            //Assert.AreEqual(1, scripter.Variables.GetVariable("resrow4").GetValue().Object);
            //Assert.AreEqual(5, scripter.Variables.GetVariable("resCol4").GetValue().Object);

            //Assert.AreEqual(2, scripter.Variables.GetVariable("resrow5").GetValue().Object);
            //Assert.AreEqual(5, scripter.Variables.GetVariable("resCol5").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow6").GetValue().Object);
            Assert.AreEqual(27, scripter.Variables.GetVariable("resCol6").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow7").GetValue().Object);
            Assert.AreEqual(52, scripter.Variables.GetVariable("resCol7").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow8").GetValue().Object);
            Assert.AreEqual(53, scripter.Variables.GetVariable("resCol8").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow9").GetValue().Object);
            Assert.AreEqual(728, scripter.Variables.GetVariable("resCol9").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow10").GetValue().Object);
            Assert.AreEqual(731, scripter.Variables.GetVariable("resCol10").GetValue().Object);

            Assert.AreEqual(5, scripter.Variables.GetVariable("resrow11").GetValue().Object);
            Assert.AreEqual(26, scripter.Variables.GetVariable("resCol11").GetValue().Object);
        }

        [Test]
        public void DoNotSearchInOtherSheetsTest()
        {
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.setsheet {SpecialChars.Text}Arkusz2{SpecialChars.Text} result {SpecialChars.Variable}res
            xlsx.find 1234 resultrow {SpecialChars.Variable}resrow resultcolumn {SpecialChars.Variable}resCol
            xlsx.find {SpecialChars.Text}abcd{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow2 resultcolumn {SpecialChars.Variable}resCol2
            xlsx.find 150 resultrow {SpecialChars.Variable}resrow3 resultcolumn {SpecialChars.Variable}resCol3
            -xlsx.find {SpecialChars.Text}160%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow4 resultcolumn {SpecialChars.Variable}resCol4
            -xlsx.find {SpecialChars.Text}100%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow5 resultcolumn {SpecialChars.Variable}resCol5
            xlsx.find {SpecialChars.Text}AA{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow6 resultcolumn {SpecialChars.Variable}resCol6
            xlsx.find {SpecialChars.Text}AZ{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow7 resultcolumn {SpecialChars.Variable}resCol7
            xlsx.find {SpecialChars.Text}BA{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow8 resultcolumn {SpecialChars.Variable}resCol8
            xlsx.find {SpecialChars.Text}AAZ{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow9 resultcolumn {SpecialChars.Variable}resCol9
            xlsx.find {SpecialChars.Text}ABC{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow10 resultcolumn {SpecialChars.Variable}resCol10
            xlsx.find {SpecialChars.Text}Z{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow11 resultcolumn {SpecialChars.Variable}resCol11

";
            scripter.Run();
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow2").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol2").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow3").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol3").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow6").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol6").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow7").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol7").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow8").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol8").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow9").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol9").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow10").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol10").GetValue().Object);

            Assert.AreEqual(-1, scripter.Variables.GetVariable("resrow11").GetValue().Object);
            Assert.AreEqual(-1, scripter.Variables.GetVariable("resCol11").GetValue().Object);
        }

        [Test]
        [Timeout(35000)]
        public void XlsxFindPercentTest()
        {
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.find {SpecialChars.Text}160%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow4 resultcolumn {SpecialChars.Variable}resCol4
            xlsx.find {SpecialChars.Text}100%{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow5 resultcolumn {SpecialChars.Variable}resCol5
";
            scripter.Run();
            Assert.AreEqual(1, scripter.Variables.GetVariable("resrow4").GetValue().Object);
            Assert.AreEqual(5, scripter.Variables.GetVariable("resCol4").GetValue().Object);

            Assert.AreEqual(2, scripter.Variables.GetVariable("resrow5").GetValue().Object);
            Assert.AreEqual(5, scripter.Variables.GetVariable("resCol5").GetValue().Object);
        }

        [Test]
        [Timeout(35000)]
        public void XlsxFindDateTest()
        {
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.find {SpecialChars.Text}21.07.2017{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow1 resultcolumn {SpecialChars.Variable}resCol1
            xlsx.find {SpecialChars.Text}22.07.2017{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow2 resultcolumn {SpecialChars.Variable}resCol2
";
            scripter.Run();
            Assert.AreEqual(1, scripter.Variables.GetVariable("resrow1").GetValue().Object);
            Assert.AreEqual(3, scripter.Variables.GetVariable("resCol2").GetValue().Object);

            Assert.AreEqual(2, scripter.Variables.GetVariable("resrow2").GetValue().Object);
            Assert.AreEqual(3, scripter.Variables.GetVariable("resCol2").GetValue().Object);
        }

        [Test]
        [Timeout(35000)]
        public void XlsxFailToFind()
        {
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.find {SpecialChars.Text}01.01.1001{SpecialChars.Text} resultrow {SpecialChars.Variable}resrow1 resultcolumn {SpecialChars.Variable}resCol1
            ";
            scripter.Run();
            Assert.AreEqual(-1, int.Parse(scripter.Variables.GetVariable("resrow1").GetValue().ToString()));
            Assert.AreEqual(-1, int.Parse(scripter.Variables.GetVariable("resCol1").GetValue().ToString()));
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
