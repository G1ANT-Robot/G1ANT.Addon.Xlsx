﻿/**
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
    [Command(Name = "xlsx.paste", Tooltip = "Pastes text from clipboard into the selected cells")]
    public class XlsxPasteCommand : Command
    {
        public XlsxPasteCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(CommandArguments arguments)
        {
            XlsxManager.CurrentXlsx.PasteFromClipboard();
        }
    }
}
