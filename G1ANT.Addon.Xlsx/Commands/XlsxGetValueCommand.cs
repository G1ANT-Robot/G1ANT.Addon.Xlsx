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
    [Command(Name = "xlsx.getvalue", Tooltip = "This command gets a value of a specified cell in an .xls(x) file")]
    public class XlsxGetValueCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxGetValueCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            object col = null;
            try
            {
                int row = arguments.Row.Value;
                if (arguments.ColIndex != null)
                    col = arguments.ColIndex.Value;
                else if (arguments.ColName != null)
                    col = arguments.ColName.Value;
                else
                    throw new ArgumentException("One of the ColIndex or ColName arguments have to be set up.");

                var result = new TextStructure(XlsxManager.CurrentXlsx.GetValue(row, col.ToString()));
                Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while getting value from specified cell. Row: {arguments.Row.Value}. Column: '{col?.ToString()}'. Message: '{ex.Message}'", ex);
            }
            
        }
    }
}
