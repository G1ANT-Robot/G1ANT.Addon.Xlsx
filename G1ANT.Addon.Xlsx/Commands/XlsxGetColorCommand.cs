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
    [Command(Name = "xlsx.getcolor", Tooltip = "Gets color of given cells to the clipboard")]
    public class XlsxGetColorCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell’s row number or row’s name")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure Column { get; set; }

            [Argument]
            public VariableStructure FontColor { get; set; } = new VariableStructure("fontcolorresult");

            [Argument]
            public VariableStructure BgColor { get; set; } = new VariableStructure("backgroundcolorresult");
        }

        public XlsxGetColorCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var colors = XlsxManager.CurrentXlsx.GetCellColor(arguments.Row.Value, arguments.Column.Value);

            Scripter.Variables.SetVariableValue(arguments.BgColor.Value, new ColorStructure(colors.Item1));
            Scripter.Variables.SetVariableValue(arguments.FontColor.Value, new ColorStructure(colors.Item2));
        }
    }
}
