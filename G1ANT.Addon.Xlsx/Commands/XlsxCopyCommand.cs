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
using G1ANT.Language;

namespace G1ANT.Addon.Xlsx
{
    [Command(Name = "xlsx.copy", Tooltip = "Copies content of given cells to the clipboard")]
    public class XlsxCopyCommand : Command
    {
        public XlsxCopyCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(CommandArguments arguments)
        {
            XlsxManager.CurrentXlsx.CopySelectionToClipboard();
        }
    }
}
