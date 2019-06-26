# xlsx.open

## Syntax

```G1ANT
xlsx.open path ⟦text⟧ accessmode ⟦text⟧ createifnotexist ⟦bool⟧
```

## Description

This command opens an .xls(x) file and activates the first sheet in the document.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes|  | Path of a file to be opened |
|`accessmode`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes|  | Can be `read` or `readwrite` |
|`createifnotexist`| [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no | false | If a file doesn’t exist, the command will create it |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the ID number of this Excel instance will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In the following script the document.xlsx is opened without the possibility to modify it (read-only mode). If the specified file doesn’t exist, it will be created by the command. The ID of this Excel instance is assigned to the `♥excelId1` variable. This ID can then be used with the [`xlsx.switch`](XlsxSwitchCommand.md) command.

```G1ANT
xlsx.open C:\Documents\document.xlsx accessmode read createifnotexist true result ♥excelId1
```

