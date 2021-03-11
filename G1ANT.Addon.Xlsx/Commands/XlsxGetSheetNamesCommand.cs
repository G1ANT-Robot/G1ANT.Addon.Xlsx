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
using System.Linq;
using G1ANT.Language;

namespace G1ANT.Addon.Xlsx.Commands
{
    [Command(Name = "xlsx.getsheetnames", Tooltip = "This command returns list of sheets in the workbook")]
    class XlsxGetSheetNamesCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public XlsxGetSheetNamesCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var sheetNames = XlsxManager.CurrentXlsx.GetSheetsNames();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new ListStructure(sheetNames.Cast<object>(), "", Scripter));
        }
    }
}
