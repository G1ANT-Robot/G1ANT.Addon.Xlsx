# xlsx.switch

## Syntax

```G1ANT
xlsx.switch id ⟦integer⟧
```

## Description

This command switches between opened .xls(x) files.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`id`| [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | ID of an opened file to switch to |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the command's result will be stored: : `true` if it succeeded, `false` if it did not |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In the following script the robot opens two .xlsx files (be sure to provide a real filepath there) and assigns their IDs to the `♥id1` and `♥id2` variables. Then it switches to the first file, reads the value of the cell A1, displays it in a dialog box and repeats the same actions on the second file:

```G1ANT
xlsx.open C:\Documents\Book1.xlsx accessmode read result ♥id1
xlsx.open C:\Documents\Book2.xlsx accessmode readwrite result ♥id2
xlsx.switch ♥id1
xlsx.getvalue row 1 colname A result ♥value1
dialog ♥value1
xlsx.switch id ♥id2
xlsx.getvalue row 1 colname A result ♥value2
dialog ♥value2
```
