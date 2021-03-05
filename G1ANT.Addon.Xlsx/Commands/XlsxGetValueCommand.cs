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
        enum ResultType
        { 
            Text,
            Structure
        }

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

            [Argument(Tooltip = "Type of the result, correct values: Text,Structure")]
            public TextStructure ResultAs { get; set; } = new TextStructure(ResultType.Text.ToString());
        }

        public XlsxGetValueCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        private Structure GetCellValue(int row, string column, ResultType resultAs)
        {
            switch (resultAs)
            {
                case ResultType.Text:
                    return new TextStructure(XlsxManager.CurrentXlsx.GetValue(row, column).ToString());
                case ResultType.Structure:
                    try
                    {
                        var val = XlsxManager.CurrentXlsx.GetValue(row, column);
                        return Scripter.Structures.CreateStructure(val, "", val?.GetType());
                    }
                    catch
                    {
                        return new TextStructure(XlsxManager.CurrentXlsx.GetValue(row, column).ToString());
                    }
                default:
                    throw new ArgumentException($"Unknown 'resultAs' argument value.");
            }
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

                ResultType resultAs = ResultType.Text;
                if (!Enum.TryParse(arguments.ResultAs.Value, true, out resultAs))
                    throw new ArgumentException($"ResultAs is not correct value. It can be one of: {string.Join(",", Enum.GetNames(typeof(ResultType)))}");

                Scripter.Variables.SetVariableValue(arguments.Result.Value, GetCellValue(row, col.ToString(), resultAs));
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while getting value from specified cell. Row: {arguments.Row.Value}. Column: '{col?.ToString()}'. Message: '{ex.Message}'", ex);
            }
            
        }
    }
}
