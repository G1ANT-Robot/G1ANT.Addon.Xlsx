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

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxGetValueTests
    {
        Scripter scripter;
        string file;

        [OneTimeSetUp]
        [Timeout(10000)]
        public void ClassInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(file));
        }
      
        [Test]
        [Timeout(10000)]
        public void XlsxGetValueDifferentTypesTest()
        {
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.getvalue row 1 colindex 1 result {SpecialChars.Variable}result1
            xlsx.getvalue row 1 colindex 2 result {SpecialChars.Variable}result2
            xlsx.getvalue row 1 colindex 4 result {SpecialChars.Variable}result3
            xlsx.getvalue row 2 colindex 7 result {SpecialChars.Variable}result4
            xlsx.getvalue row 5 colindex 27 result {SpecialChars.Variable}result6
            xlsx.getvalue row 5 colindex 52 result {SpecialChars.Variable}result7
            xlsx.getvalue row 5 colindex 53 result {SpecialChars.Variable}result8
            xlsx.getvalue row 5 colindex 728 result {SpecialChars.Variable}result9
            xlsx.getvalue row 5 colindex 731 result {SpecialChars.Variable}result10
            xlsx.getvalue row 5 colindex 26 result {SpecialChars.Variable}result11
            xlsx.getvalue row 22 colindex 38 result {SpecialChars.Variable}result12
";
            scripter.Run();
            Assert.AreEqual("1234", scripter.Variables.GetVariable("result1").GetValue().Object);
            Assert.AreEqual("abcd", scripter.Variables.GetVariable("result2").GetValue().Object);
            Assert.AreEqual("150", scripter.Variables.GetVariable("result3").GetValue().Object);
            
            Assert.AreEqual("AA", scripter.Variables.GetVariable("result6").GetValue().Object);
            Assert.AreEqual("AZ", scripter.Variables.GetVariable("result7").GetValue().Object);
            Assert.AreEqual("BA", scripter.Variables.GetVariable("result8").GetValue().Object);
            Assert.AreEqual("AAZ", scripter.Variables.GetVariable("result9").GetValue().Object);
            Assert.AreEqual("ABC", scripter.Variables.GetVariable("result10").GetValue().Object);
            Assert.AreEqual("Z", scripter.Variables.GetVariable("result11").GetValue().Object);

            Assert.IsTrue(string.IsNullOrEmpty(scripter.Variables.GetVariable("result12").GetValue().Object.ToString()));
        }

        [Test]
        [Timeout(10000)]
        public void XlsxGetValuePercentTest()
        {
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.getvalue row 1 colindex 5 result {SpecialChars.Variable}result1
            xlsx.getvalue row 2 colindex 5 result {SpecialChars.Variable}result2";
            scripter.Run();

            Assert.AreEqual("160%", scripter.Variables.GetVariable("result1").GetValue().Object);
            Assert.AreEqual("100%", scripter.Variables.GetVariable("result2").GetValue().Object);
        }

        [Test]
        [Timeout(10000)]
        public void XlsxGetValueFloatTest()
        {
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.getvalue row 2 colindex 7 result {SpecialChars.Variable}result1
            ";
            scripter.Run();
            Assert.AreEqual(12.345f, scripter.Variables.GetVariable("result4").GetValue().Object);
        }
        [OneTimeTearDown]
        [Timeout(10000)]
        public void ClassCleanUp()
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }
}