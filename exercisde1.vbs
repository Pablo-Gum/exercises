option Explicit
Dim hours,rate 'Input
Dim salary 'output

hours = CDbl(InputBox("Hours worked"))
rate = CDbl(InputBox("rate of pay"))

salary = hours * rate

msgBox("your salary is "& salary)
