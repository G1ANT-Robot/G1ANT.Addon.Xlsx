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
    [Command(Name = "xlsx.setcolumnwidth",Tooltip = "This command sets a width of a specified column in an .xls(x) file")]
    public class XlsxSetColumnWidthCommand : Command
    {
        public  class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Column name")]
            public TextStructure ColName { get; set; }

            [Argument(Tooltip = "Column width", Required = false)]
            public IntegerStructure Width{ get; set; }

            [Argument(Tooltip = "Adjust width to content", Required = false)]
            public BooleanStructure AdjustToContents { get; set; } = new BooleanStructure(false);
        }
        public XlsxSetColumnWidthCommand(AbstractScripter scripter) : base(scripter)
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
                    throw new ArgumentException("One of the ColIndex or ColName arguments must be provided.");

                XlsxManager.CurrentXlsx.SetColumnWidth(col.ToString(), arguments.Width?.Value, arguments.AdjustToContents.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while setting column width.", ex);
            }
        }
    }
}
