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
    [Command(Name = "xlsx.switch", Tooltip = "This command allows to switch between open .xlsx files.")]
    public class XlsxSwitchCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true)]
            public IntegerStructure Id { get; set; } = new IntegerStructure(0);

            [Argument]
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
