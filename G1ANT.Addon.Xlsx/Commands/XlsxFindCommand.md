# xlsx.find

## Syntax

```G1ANT
xlsx.find value ⟦text⟧ resultcolumn ⟦variable⟧ resultrow ⟦variable⟧
```

## Description

This command finds an address of a cell where a specified value is stored.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`value`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes  | | Value to be searched for |
| `resultcolumn` | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥resultcolumn`                                           | Name of a variable where the command's result (column index) will be stored |
| `resultrow` | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥resultrow`                                         | Name of a variable where the command's result (row number) will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

This simple script opens an Excel workbook (be sure to provide a real filepath) and searches for “*aaa*” value. If it doesn’t find any match, an error message appears. If there’s a cell containing the searched value, its coordinates are displayed in a dialog box:

```G1ANT
xlsx.open C:\Tests\Book1.xlsx
xlsx.find aaa errormessage ‴Value not found‴
dialog ‴Value found in the cell: column ♥resultcolumn, row ♥resultrow‴
```

