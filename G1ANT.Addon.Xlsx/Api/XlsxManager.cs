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
using System.Linq;


namespace G1ANT.Addon.Xlsx
{
    public static class XlsxManager
    {
        private static List<XlsxWrapper> launchedXlsx = new List<XlsxWrapper>();

        public static XlsxWrapper CurrentXlsx { get; private set; }

        public static int getFirstId()
        {
            return launchedXlsx.First().Id;
        }

        private static int GetNextId()
        {
            return launchedXlsx.Count() > 0 ? launchedXlsx.Max(x => x.Id) + 1 : 0;
        }

        public static XlsxWrapper AddXlsx()
        {
            int assignedId = GetNextId();
            XlsxWrapper wrapper = new XlsxWrapper(assignedId);
            launchedXlsx.Add(wrapper);
            CurrentXlsx = wrapper;
            return wrapper;
        }

        public static bool Open(string filePath)
        {
            if (CurrentXlsx.Open(filePath)) return true;
            else return false;
        }

        public static int CountRows()
        {
            return CurrentXlsx.CountRows();
        }

        public static bool SwitchXlsx(int id)
        {
            var tmpXlsx = launchedXlsx.Where(x => x.Id == id).FirstOrDefault();
            CurrentXlsx = tmpXlsx ?? CurrentXlsx;
            return tmpXlsx != null;
        }

        public static void Remove(int id)
        {
            XlsxWrapper toRemove = launchedXlsx.Where(x => x.Id == id).FirstOrDefault();
            Remove(toRemove);
        }

        public static void Remove(XlsxWrapper wrapper)
        {
            if (wrapper != null)
            {
                wrapper.Close();
                if (launchedXlsx.Contains(wrapper))
                {
                    launchedXlsx.Remove(wrapper);
                }
                CurrentXlsx = launchedXlsx.FirstOrDefault();
            }
        }
    }
}
