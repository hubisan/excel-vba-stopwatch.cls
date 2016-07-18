# Stopwatch to measure elapsed time in milliseconds or seconds

Start a timer and get the elapsed time in milliseconds or seconds. Mainly using the class to measure and compare the speed of macros.

## Installation

- Download repository: Clone or [download](../../archive/master.zip) the repository and import the file Stopwatch.cls into your workbook to use the class.
- Or download the class file only: You can also download the class file "Stopwatch.cls" directly by right-clicking and selecting "save link as&hellip;".
- Or copy the code: Copy the code starting at line 10 `#If VBA7` and paste into a new class in your workbook.

### Usage

Methods and properties:

- `.start`: start or restart the timer
- `.Elapsed_ms`: get the elapsed time in milliseconds (like 1004) As Double
- `.Elapsed_sec(Optional number_of_digits_after_decimal As Integer = 3)`: get the elapsed time in seconds as Double.  
  You can round to a number of digits after the decimal by providing the optional parameter, the default is 3 = millliseconds.

Example usage:

```vb
Sub example()
    Dim x As New Stopwatch
    x.start
    'do your stuff here
    'in ms. Output: 1004
    Debug.Print x.Elapsed_ms  'output like 1004
    'in seconds rounded to 3 digits after decimal. Output: 1.004
    Debug.Print x.Elapsed_sec
    'in seconds rounded to 0 digits after decimal. Output: 1
    Debug.Print x.Elapsed_sec(0)
End Sub
```
## Contributing

If you find a bug, please create a new issue. Pull requests are also welcome.

## Contributors

- [Daniel Hubmann](https://github.com/hubisan) (Author)

## License

Copyright (c) 2016 Daniel Hubmann. Licensed under [MIT](LICENSE).
