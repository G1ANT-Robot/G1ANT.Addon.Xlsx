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
using G1ANT.Addon.Xlsx.Api;
using G1ANT.Language;

namespace G1ANT.Addon.Xlsx
{
    [Command(Name = "xlsx.select", Tooltip = "Selects continuous set of cells defined by two opposite vertices")]
    public class XlsxSelectCommand : Command
    {
        public class CopyArguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "First cell's row number")]
            public IntegerStructure Row1 { get; set; }

            [Argument(Required = true, Tooltip = "First cell's column identifier")]
            public TextStructure Column1 { get; set; }

            [Argument(Required = true, Tooltip = "Second cell's row number")]
            public IntegerStructure Row2 { get; set; }

            [Argument(Required = true, Tooltip = "Second cell's column identifier")]
            public TextStructure Column2 { get; set; }
        }

        public XlsxSelectCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(CopyArguments arguments)
        {
            XlsxManager.CurrentXlsx.SelectRange(
                new CellRef(XlsxManager.CurrentXlsx.ActiveSheetId, XlsxManager.CurrentXlsx.FormatInput(arguments.Column1.Value, arguments.Row1.Value)),
                new CellRef(XlsxManager.CurrentXlsx.ActiveSheetId, XlsxManager.CurrentXlsx.FormatInput(arguments.Column2.Value, arguments.Row2.Value))
                );
        }
    }
}
