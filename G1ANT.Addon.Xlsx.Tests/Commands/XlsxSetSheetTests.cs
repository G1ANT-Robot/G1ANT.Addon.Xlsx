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
using System.IO;

using G1ANT.Engine;
using NUnit.Framework;
using System.Reflection;
using G1ANT.Language;
using G1ANT.Addon.Xlsx.Tests.Properties;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxSetSheetTests
    {
        string file;
        Scripter scripter;
        [OneTimeSetUp]
        [Timeout(20000)]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(file));
        }

        [Test]
        [Timeout(20000)]
        public void XlsxSetSheetDefault()
        {
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.setsheet result {SpecialChars.Variable}res
            ";
            scripter.Run();
            Assert.IsTrue(scripter.Variables.GetVariableValue<bool>("res"));
        }

        [Test]
        [Timeout(20000)]
        public void XlsxSetSheetCustom()
        {
            scripter.Text = $@"
            xlsx.open {SpecialChars.Variable}xlsPath result {SpecialChars.Variable}id
            xlsx.setsheet {SpecialChars.Text}Arkusz2{SpecialChars.Text} result {SpecialChars.Variable}res
            ";
            scripter.Run();
            Assert.IsTrue(scripter.Variables.GetVariableValue<bool>("res"));
        }

        [Test]
        [Timeout(20000)]
        public void SetNotExistingSheet()
        {
            scripter.Text = $"xlsx.setsheet a!@#$poq098239 result {SpecialChars.Variable}res";
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsFalse(scripter.Variables.GetVariableValue<bool>("res"));
            Assert.IsInstanceOf<ArgumentOutOfRangeException>(exception.GetBaseException());
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
