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
    [Command(Name = "xlsx.setcolor", Tooltip = "Sets the color of given cell")]
    public class XlsxSetColorCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell’s row number or row’s name")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure Column { get; set; }

            [Argument]
            public ColorStructure FontColor { get; set; }

            [Argument]
            public ColorStructure BgColor { get; set; }
        }

        public XlsxSetColorCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (arguments.BgColor != null)
            {
                XlsxManager.CurrentXlsx.SetCellBackgroundColor(
                    arguments.Row.Value, arguments.Column.Value,
                    arguments.BgColor.Value);
            }

            if (arguments.FontColor != null)
            {
                XlsxManager.CurrentXlsx.SetCellFontColor(
                    arguments.Row.Value, arguments.Column.Value,
                    arguments.FontColor.Value);
            }
        }
    }
}
