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
    [Command(Name = "xlsx.setvalue",Tooltip = "This command sets a value of a specified cell in an .xls(x) file")]
    public class XlsxSetValueCommand : Command
    {
        public  class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Value to be set")]
            public TextStructure Value { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxSetValueCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            object col = null;
            try
            {
                if (arguments.ColIndex != null)
                    col = arguments.ColIndex.Value;
                else if (arguments.ColName != null)
                    col = arguments.ColName.Value;
                else
                    throw new ArgumentException("One of the ColIndex or ColName arguments have to be set up.");
                XlsxManager.CurrentXlsx.SetValue(arguments.Row.Value, col.ToString(), arguments.Value.Value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(true));
            }
            catch (Exception ex)
            {
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(false));
                throw new ApplicationException($"Problem occured while setting value. Row: '{arguments.Row.Value}', Col: '{col}', Val: '{arguments.Value.Value}'", ex);
            }
        }
    }
}
