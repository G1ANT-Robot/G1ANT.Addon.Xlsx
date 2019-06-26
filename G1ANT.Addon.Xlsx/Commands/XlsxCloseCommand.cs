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
    [Command(Name = "xlsx.close",Tooltip = "This command allows to save changes and close .xlsx file.")]
    public class XlsxCloseCommand :Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "ID of file to close. If not set, will close file opened as first")]
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
