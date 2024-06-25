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
using System.Linq;
using System.Drawing;
using G1ANT.Language;

namespace G1ANT.Addon.Xlsx
{
    [Command(Name = "xlsx.findall", Tooltip = "This command finds all addressess of cells where a specified value is stored")]
    public class XlsxFindAllCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Value to be searched for")]
            public TextStructure Value { get; set; } = new TextStructure("value");

            [Argument(Required = true, Tooltip = "If true search the value only in selection")]
            public BooleanStructure InSelection { get; set; } = new BooleanStructure(false);
            
            [Argument(Required = true, Tooltip = "Indicates that the string comparison must ignore case")]
            public BooleanStructure IgnoreCase { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Name of a variable where the command's result (list of points) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public XlsxFindAllCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            ListStructure result = new ListStructure(null, "", Scripter);
            var searchResult = XlsxManager.CurrentXlsx.Find(arguments.Value.Value, arguments.InSelection.Value, arguments.IgnoreCase.Value);
            if (searchResult != null)
            {
                var list = searchResult.
                    Select(x => new PointStructure(new Point(x.ColumnNumber, x.RowNumber))).ToList<object>();
                result = new ListStructure(list, "", Scripter);
            }
            Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
        }
    }
}
