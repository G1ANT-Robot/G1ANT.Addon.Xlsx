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

namespace G1ANT.Addon.Xlsx
{
    [Command(Name = "xlsx.countrows",Tooltip = "This command counts rows in an open .xls(x) file")]
    public class XlsxCountRowsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxCountRowsCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public  void Execute(Arguments arguments)
        {
            try
            {
                int res = XlsxManager.CurrentXlsx.CountRows();
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(res));
            }
            catch
            {
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(-1));
            }
        }

      
    }
}
