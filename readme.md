# Stopwatch to measure elapsed time in milliseconds or seconds

Excel VBA class to Start a timer and get the elapsed time in milliseconds or seconds.

## Installation

- **Download the workbook**: You can download [stopwatch.xlsm](../../raw/master/stopwatch.xlsm). This workbook inlcudes the class and a module with two examples on how to use the class.
- **Or copy the code**: Copy the class code starting at line 10 `#If VBA7` from [raw](../../raw/master/Stopwatch.cls) and paste into a new class in your workbook.
- **Or get the repository**: Clone or [download](../../archive/master.zip) the repository and import the file Stopwatch.cls into your workbook to use the class.

### Usage

Methods and properties:

- `.start`: start or restart the timer
- `.Elapsed_ms`: get the elapsed time in milliseconds (like 1004) As Double
- `.Elapsed_sec(Optional number_of_digits_after_decimal As Integer = 3)`: get the elapsed time in seconds as Double.  
  You can round to a number of digits after the decimal by providing the optional parameter, the default is 3 = millliseconds.
- `.stop_it`: stops the timer, rarely using this myself.

Example usage:

```vb
Sub example()
    Dim x As New Stopwatch
    x.start
    'do your stuff here
    'in ms. Output: 1004
    Debug.Print x.Elapsed_ms
    'in seconds rounded to 3 digits after decimal. Output: 1.004
    Debug.Print x.Elapsed_sec
    'in seconds rounded to 0 digits after decimal. Output: 1
    Debug.Print x.Elapsed_sec(0)
End Sub
```

## Remark for the workbook

The class module gets automatically exported when saving the Excel file. To stop this behavior change the constant in the module "autoopen" to "False":
```vb
Private Const is_in_development As Boolean = False
```

## Contributing

If you find a bug, please create a new issue. Pull requests are also welcome.

## Contributors

- [Daniel Hubmann](https://github.com/hubisan) (Author)

## License

Copyright (c) 2016 Daniel Hubmann. Licensed under [MIT](LICENSE).
