/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System.Collections.Generic;
using System.IO;

using G1ANT.Engine;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace G1ANT.Addon.Xlsx.Tests
{
    [TestClass]
    public class Initializer
    {
        private static List<string> filesToDelete = new List<string>();

        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            UnPackResources();
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            CleanupResources();
        }

        public static string TestWorkBookPath { get; private set; }

        public static string EmpyWorkBookPath { get; private set; }

        //TODO public static string DirectoryPath { get; private set; } = SettingsContainer.Instance.Directories[Infrastructure.Source.UserDocsDir].FullName;//TODO WhatIsInfrastructure??

        private static void UnPackResources()
        {
            throw new NotImplementedException();
            if (File.Exists(TestWorkBookPath))
            {
                File.Delete(TestWorkBookPath);
            }
            filesToDelete.Add(TestWorkBookPath);
            File.WriteAllBytes(TestWorkBookPath, Properties.Resources.XlsTestWorkbook);

            if (File.Exists(EmpyWorkBookPath))
            {
                File.Delete(EmpyWorkBookPath);
            }
            filesToDelete.Add(EmpyWorkBookPath);
            File.WriteAllBytes(EmpyWorkBookPath, Properties.Resources.EmptyWorkbook);
        }

        private static void CleanupResources()
        {
            if (filesToDelete != null)
            {
                foreach (string path in filesToDelete)
                {
                    File.Delete(path);
                }
            }
        }

        public static bool AreEqual(byte[] a, byte[] b)
        {
            if (a.Length != b.Length)
            {
                return false;
            }

            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                {
                    return false;
                }
            }

            return true;
        }
    }
}
