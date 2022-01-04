# Gradebook

An excel workbook for recording grades.

## How to use

Start from the `MainPage`. Should be intuitive enough.

This workbook uses VBA macros. You may need to enable macros (by clicking "Enable Content") first.

If you do not enable macros, you should still be able to use the workbook, but some features of the workbook would be limited, e.g. you can only edit existing records.

## Build

The current build process is as follows:

1. Build the (macro-free) skeleton `xlsm` with `buildSkeleton.bat`
2. Open the VBA IDE in the workbook in `MS Excel`
3. Import `src` using the `Sync Project -> Update Components` command in the [`Rubberduck` addin](https://github.com/rubberduck-vba/Rubberduck) 
4. Save the file

Doing so should give you a functioning `xlsm` workbook

## FAQ

### Why excel?

It is popular enough that

1. `MS Excel` is already on most machines
2. even non tech people know how to use it

### Why macros?

Some functions cannot (easily) be implemented with excel formula and using macros is the solution

## TODO

* ~~figure out how to version control `xlsx`~~
  * currently using `Rubberduck`
* allow partial ID input in work record form
* allow entering student name in work record form
* figure out how to get the student list automatically
* figure out how to neutralise `printerSettings1.bin`

