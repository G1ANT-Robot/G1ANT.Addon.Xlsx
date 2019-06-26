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
    [Command(Name = "xlsx.open", Tooltip = "This command allows to open .xlsx files, and set the first sheet in the document as active.")]
    public class XlsxOpenCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true)]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = false)]
            public TextStructure AccessMode { get; set; } = new TextStructure(string.Empty);

            [Argument]
            public BooleanStructure CreateIfNotExist { get; set; } = new BooleanStructure(false);

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxOpenCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            string returVariableName = arguments.Result.Value;
            var xlsWraper = XlsxManager.AddXlsx();
            try
            {
                if (arguments.CreateIfNotExist.Value)
                {
                    if (System.IO.File.Exists(arguments.Path.Value) == false)
                        xlsWraper.Create(arguments.Path.Value);
                }
                if (xlsWraper.Open(arguments.Path.Value, arguments.AccessMode.Value))
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(xlsWraper.Id));
                    OnScriptEnd = () =>
                    {
                        XlsxManager.Remove(xlsWraper);
                    };
                }
                else
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(-1));
                }
            }
            catch (Exception ex)
            {
                if (xlsWraper != null)
                {
                    XlsxManager.Remove(xlsWraper);
                }
                throw ex;
            }
        }
    }
}
