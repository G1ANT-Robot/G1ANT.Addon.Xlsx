# xlsx.setvalue

## Syntax

```G1ANT
xlsx.setvalue value ⟦text⟧ row ⟦integer⟧ colname ⟦text⟧
```

or

```G1ANT
xlsx.setvalue value ⟦text⟧ row ⟦integer⟧ colindex ⟦integer⟧
```

## Description

This command sets a value of a specified cell in an .xls(x) file.

| Argument                | Type                                                         | Required | Default Value                                                | Description                                                  |
| ----------------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `value`                 | [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | Value to be set                                              |
| `row`                   | [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes      |                                                              | Cell's row number                                            |
| `colindex` or `colname` | [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) or [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | `colindex`: cell's column number, `colname`: cell's column name |
| `result`                | [variable](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                    | Name of a variable where the command's result will be stored |
| `if`                    | [bool](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`               | [timespan](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`             | [procedure](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`             | [label](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage`          | [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`           | [variable](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](https://manual.g1ant.com/link/G1ANT.Manual/appendices/common-arguments.md) page.

## Example

The following script opens a workbook (if you don’t specify a real filepath, a new file will be created), enters *123* into the cell A1 and closes the file, saving the changes:

```G1ANT
xlsx.open C:\Documents\TestBook.xlsx createifnotexist true
xlsx.setvalue value 123 row 1 colname a
xlsx.close
```

