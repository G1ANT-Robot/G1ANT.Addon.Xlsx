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
    [Command(Name = "xlsx.setsheet", Tooltip = "This command sets a specified sheet as active")]
    public class XlsxSetSheetCommand : Command
    {
        public  class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a sheet to be set as active. If it’s not specified, the robot will activate the first sheet in the file")]
            public TextStructure Name { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored: `true` if it succeeded, `false` if it did not")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxSetSheetCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public  void Execute(Arguments arguments)
        {
            if (arguments.Name.Value == string.Empty)
            {
                try
                {
                    XlsxManager.CurrentXlsx.ActivateSheet(XlsxManager.CurrentXlsx.GetSheetsNames()[0]);
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(true));
                }
                catch
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(false));
                    throw new InvalidOperationException("Workbook do not have any sheet");
                }
            }
            else
            {
                try
                {
                    XlsxManager.CurrentXlsx.ActivateSheet(arguments.Name.Value);
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(true));
                }
                catch
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(false));
                    throw new ArgumentOutOfRangeException(nameof(arguments.Name), arguments.Name.Value, $"Workbook do not have '{arguments.Name.Value} sheet");
                }
            }
        }
    }
}
