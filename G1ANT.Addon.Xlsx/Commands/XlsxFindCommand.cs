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
    [Command(Name = "xlsx.find", Tooltip = "This command finds an address of a cell where a specified value is stored")]
    public class XlsxFindCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Value to be searched for")]
            public TextStructure Value { get; set; } = new TextStructure("value");

            [Argument(Tooltip = "Name of a variable where the command's result (column index) will be stored")]
            public VariableStructure ResultColumn { get; set; } = new VariableStructure("resultcolumn");
            [Argument(Tooltip = "Name of a variable where the command's result (row number) will be stored")]
            public VariableStructure ResultRow { get; set; } = new VariableStructure("resultrow");
        }
        public XlsxFindCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            string position = XlsxManager.CurrentXlsx.Find(arguments.Value.Value);

            if (position != null)
            {
                int[] columRowPair = XlsxManager.CurrentXlsx.FormatInput(position);
                Scripter.Variables.SetVariableValue(arguments.ResultColumn.Value, new Language.IntegerStructure(columRowPair[0]));
                Scripter.Variables.SetVariableValue(arguments.ResultRow.Value, new Language.IntegerStructure(columRowPair[1]));
            }
            else
            {
                Scripter.Variables.SetVariableValue(arguments.ResultColumn.Value, new Language.IntegerStructure("-1"));
                Scripter.Variables.SetVariableValue(arguments.ResultRow.Value, new Language.IntegerStructure("-1"));
            }

        }
    }
}
