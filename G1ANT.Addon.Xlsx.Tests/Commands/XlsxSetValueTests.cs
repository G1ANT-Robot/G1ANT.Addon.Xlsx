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
using G1ANT.Engine;


using NUnit.Framework;
using System.Reflection;
using G1ANT.Language;
using G1ANT.Addon.Xlsx.Tests.Properties;
using System.IO;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxSetValueTests
    {
        string file;
        Scripter scripter;
        
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
            scripter.Text = $@"xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id";
            scripter.Run();
        }

        [Test]
        [Timeout(10000)]
        public void XlsxSetValueInd()
        {
            string checkVal1 = "123";
            string checkVal2 = "test";
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.setvalue value {SpecialChars.Text}{checkVal1}{SpecialChars.Text} row 3 colindex 6 result {SpecialChars.Variable}res1
            xlsx.setvalue value {SpecialChars.Text}{checkVal2}{SpecialChars.Text} row 4 colindex 6 result {SpecialChars.Variable}res2
            xlsx.getvalue row 3 colindex 6 result {SpecialChars.Variable}result1
            xlsx.getvalue row 4 colindex 6 result {SpecialChars.Variable}result2
            ";
            scripter.Run();
            Assert.AreNotEqual(false, scripter.Variables.GetVariableValue<bool>("res1"));
            Assert.AreNotEqual(false, scripter.Variables.GetVariableValue<bool>("res2"));
            Assert.AreEqual(checkVal2, scripter.Variables.GetVariableValue<string>("result2"));
            Assert.AreEqual(checkVal1, scripter.Variables.GetVariable("result1").GetValue().Object);
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.setvalue value {SpecialChars.Text}{SpecialChars.Text} row 3 colindex 6 result {SpecialChars.Variable}res1
            xlsx.setvalue value {SpecialChars.Text}{SpecialChars.Text} row 4 colindex 6 result {SpecialChars.Variable}res2
            ";
            scripter.Run();
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
