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
using G1ANT.Addon.Xlsx.Tests.Properties;
using G1ANT.Language;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestFixture]
    public class XlsxSwitchTests
    {

        string file;
        string file2;
        Scripter scripter;
        private static int filesCount = 5;
        private static string[] filePaths = new string[filesCount];
        private static string filePrefix ;

        [OneTimeSetUp]
        [Timeout(20000)]
        public void TestInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
        }

        [SetUp]
        public void testinit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            scripter = new Scripter();
scripter.InitVariables.Clear();
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            file2 = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.EmptyWorkbook), "xlsx");
            filePrefix = file;

            for (int i = 0; i < filesCount; i++)
            {
                filePaths[i] = $"{filePrefix}{i}";
                if (File.Exists(filePaths[i]) == false)
                {
                    File.Copy(file, filePaths[i]);
                }
                else
                {
                    filePaths[i] = null;
                    Assert.Inconclusive($"File '{filePrefix}{i}' exists");
                }
            }
        }
        [TearDown]
        [Timeout(20000)]
        public void TestCleanup()
        {

            foreach (string path in filePaths)
            {
                if (path != null)
                {
                    try
                    {
                        scripter.RunLine("xlsx.close");
                    }
                    catch { }
                    File.Delete(path);
                }
            }
        }

        [Test]
        [Timeout(20000)]
        public void XlsxSwitchTest()
        {
            int[] xlsIds = new int[filesCount];

            for (int i = 0; i < filesCount; i++)
            {
                scripter.RunLine($"xlsx.open {SpecialChars.Text}{filePaths[i]}{SpecialChars.Text} result {SpecialChars.Variable}id");
                xlsIds[i] = scripter.Variables.GetVariableValue<int>("id");
            }

            Random randomGenerator = new Random();

            int testsCount = 10;
            for (int i = 0; i < testsCount; i++)
            {
                int id = randomGenerator.Next(filesCount);

                scripter.RunLine($"xlsx.switch {xlsIds[id]} result {SpecialChars.Variable}hasSwitched");
                Assert.IsTrue(scripter.Variables.GetVariableValue<bool>("hasSwitched"));
            }

            for (int i = 0; i < filesCount; i++)
            {
                scripter.RunLine($"xlsx.close {xlsIds[i]}");
            }
        }

        [Test]
        [Timeout(20000)]
        public void FileChangeTest()
        {
            int[] xlsIds = new int[filesCount];
            byte[] fileBytes;

            for (int i = 0; i < filesCount; i++)
            {
                fileBytes = File.ReadAllBytes(filePaths[i]);
                if (Initializer.AreEqual(Properties.Resources.XlsTestWorkbook, fileBytes) == false)
                {
                    Assert.Inconclusive($"File '{filePaths[i]}' is different than original before editting");
                }
                scripter.RunLine($"xlsx.open {SpecialChars.Text}{filePaths[i]}{SpecialChars.Text} result {SpecialChars.Variable}id");
                xlsIds[i] = scripter.Variables.GetVariableValue<int>("id");
            }

            Random randomGenerator = new Random();

            int id = -1;

            for (int i = 0; i < filesCount; i++)
            {
                id = randomGenerator.Next(filesCount);
                scripter.RunLine($"xlsx.switch id {id}");
                scripter.RunLine($"xlsx.setvalue {SpecialChars.Text}{"some value"}{SpecialChars.Text} row 1 colindex 1");
                scripter.RunLine($"xlsx.close");

                fileBytes = File.ReadAllBytes(filePaths[i]);
                Assert.IsFalse(Initializer.AreEqual(Properties.Resources.XlsTestWorkbook, fileBytes));
            }
        }
    }
}
