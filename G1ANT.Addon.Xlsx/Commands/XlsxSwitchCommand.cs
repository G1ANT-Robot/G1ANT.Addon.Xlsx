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
    [Command(Name = "xlsx.switch", Tooltip = "This command switches between opened .xls(x) files")]
    public class XlsxSwitchCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "ID of an opened file to switch to")]
            public IntegerStructure Id { get; set; } = new IntegerStructure(0);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored: `true` if it succeeded, `false` if it did not")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxSwitchCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                int id = arguments.Id.Value;
                bool result = XlsxManager.SwitchXlsx(id);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(result));
            }
            catch
            {
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(false));
                throw new ApplicationException("Specified Xlsx not existing");
            }
        }
    }
}
