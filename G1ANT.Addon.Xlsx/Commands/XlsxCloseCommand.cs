/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Language;
using System;

namespace G1ANT.Addon.Xlsx
{
    [Command(Name = "xlsx.close",Tooltip = "This command saves changes to an .xls(x) file and closes it")]
    public class XlsxCloseCommand :Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "ID of a file to close. If not set, the first opened file will be closed")]
            public IntegerStructure Id { get; set; }
        }
        public XlsxCloseCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            int ID;
            if (arguments.Id == null)
            {
                ID = XlsxManager.getFirstId();
            }
            else
            {
                ID = arguments.Id.Value;
            }
            XlsxManager.Remove(ID);
        }
    }
}
