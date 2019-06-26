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
    [Command(Name = "xlsx.setsheet", Tooltip = "This command allows to set active sheet to work with.")]
    public class XlsxSetSheetCommand : Command
    {
        public  class Arguments : CommandArguments
        {
            [Argument]
            public TextStructure Name { get; set; } = new TextStructure(string.Empty);

            [Argument]
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
