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
    public class XlsxCloseTests
    {
        //TODO 
        // Another test to add=> after xlsx.open . RUN() should try run another line and throw "null ref exeception" => scripter should automatically release all file handles after finishing running script.
        string file;
        string file2;
        static int filesCount = 5;
        Scripter scripter;
        static string[] filePaths = new string[filesCount];
        string filePrefix;

        [OneTimeSetUp]
        public void Initialize()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            file = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.XlsTestWorkbook), "xlsx");
            file2 = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.EmptyWorkbook), "xlsx");
            filePrefix = file;
            scripter = new Scripter();
            scripter.InitVariables.Clear();
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
        [SetUp]
        public void testinit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Xlsx.dll");
            scripter = new Scripter();
            scripter.InitVariables.Clear();
        }
        [Test]
        [Timeout(20000)]
        public void CloseTest()
        {
            int[] xlsIds = new int[filesCount];

            for (int i = 0; i < filesCount; i++)
            {
                scripter.RunLine($"xlsx.open {SpecialChars.Text}{filePaths[i]}{SpecialChars.Text}");
                xlsIds[i] = scripter.Variables.GetVariableValue<int>("result");
            }

            bool gotAccessToopenedfile = true;
            FileStream openedFile = null;
            try
            {
                openedFile = File.Open(filePaths[0], FileMode.Open, FileAccess.ReadWrite);
            }
            catch
            {
                gotAccessToopenedfile = false;
            }
            finally
            {
                if (openedFile != null)
                {
                    openedFile.Close();
                    openedFile = null;
                }
            }
            Assert.IsFalse(gotAccessToopenedfile, $"Access aquired to file '{filePaths[0]}' despaite it's openede by xlsx command");

            scripter.RunLine($"xlsx.close id {xlsIds[0]}");
            gotAccessToopenedfile = true;
            openedFile = null;
            try
            {
                openedFile = File.Open(filePaths[0], FileMode.Open, FileAccess.ReadWrite);
            }
            catch (Exception)
            {
                gotAccessToopenedfile = false;
            }
            finally
            {
                if (openedFile != null)
                {
                    openedFile.Close();
                    openedFile = null;
                }
            }
            Assert.IsTrue(gotAccessToopenedfile, $"Access do not aquired to file '{filePaths[0]}' despaite it's closed by xlsx.close command");

            for (int i = 1; i < filesCount; i++)
            {
                scripter.RunLine($"xlsx.close");
            }

            for (int i = 1; i < filesCount; i++)
            {
                gotAccessToopenedfile = true;
                openedFile = null;
                try
                {
                    openedFile = File.Open(filePaths[i], FileMode.Open, FileAccess.ReadWrite);
                }
                catch (Exception)
                {
                    gotAccessToopenedfile = false;
                }
                finally
                {
                    if (openedFile != null)
                    {
                        openedFile.Close();
                        openedFile = null;
                    }
                }
                Assert.IsTrue(gotAccessToopenedfile, $"Access do not aquired to file '{filePaths[i]}' despaite it's closed by xlsx.close command");
            }
        }

        [Test]
        [Timeout(20000)]
        public void CloseWithoutOpenTest()
        {
            scripter.Text = "xlsx.close";
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<InvalidOperationException>(exception.GetBaseException());
        }


        [TearDown]
        [Timeout(20000)]
        public void TestCleanup()
        {
            foreach (string path in filePaths)
            {
                if (path != null && File.Exists(path) != false)
                {
                    File.Delete(path);
                }
            }
        }
    }
}
