Option Explicit

Dim rateOfPay,hoursWorked,grossPay,excessHours

const normalWorkingHours = 40

rateOfPay = CDbl(InputBox("Enter hourly pay:"))
hoursWorked = CInt(InputBox("Enter the number of hours you worked:"))

if hoursWorked >0 and hoursWorked <=normalWorkingHours then
grossPay = rateOfPay * normalWorkingHours
MsgBox(grossPay)
elseif hoursWorked > normalWorkingHours then
excessHours = hoursWorked - normalWorkingHours
grossPay = (rateOfPay * normalWorkingHours) + (excessHours * 1.5 * rateOfPay)
MsgBox(grossPay)
elseif hoursWorked = 0 then
grossPay = 0
MsgBox(grossPay)
else MsgBox("Incorrect input")
end if
