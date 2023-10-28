This is a simple program that allows the user to write out two cells (in the format e.g. E5, B3 AB10) and then perform any operations in the space designated by these cells.

In the given program, each cell in the space is colored based on the given interval values

Program asks user to input to paths
 
<0-30) --> red #F2220F 
<30,50) --> yellow #F2A007
<50, 75) --> green #05F26C
<75, 90) --> blue #0597F2
<90, 100) --> purple #B225D9

/////////////////////////////////////////////////////////////////
Note:

Excel file can't be open during running program

/////////////////////////////////////////////////////////////////
 
Things to improve:

*Cloning more cloning files if there is already copy of original file 
 
*The program crashes due to not limiting the memory, e.g. with a conditional instruction

Unhandled exception at 0x00007FFD1DB74FFC in Project3.exe: Microsoft C++ exception: std::length_error at memory location 0x0000001D09AFECD0.

*The program crashes due to incorrect user input(e.g. !5)

*Distinguishing between worksheet and woorkbook (e.g. create worksheet in workbook)
 
*File can be open during running program

/////////////////////////////////////////////////////////////////

Things improved:
 
*The program does not crash when trying to color a given cell
*The program clones the downloaded file.xlsx file so that the original file retains its values
*Allow user to choose excel file location by entering its path
