Dim paymentType
paymentType = LCase(InputBox("Enter payment type"))

Select Case paymentType
Case "cash"
MsgBox("You are going to pay cash")
Case "visa"
MsgBox("You are goign to pay with Visa")
Case "amex"
MsgBox("You are going to pay with American Express")
Case Else
MsgBox("Invalid payment method")
end Select

'APPROACH TWO
dim intAmtCopy
dim dblAmount
dim cnt

For cnt=1 to 3
intAmtCopy = CDbl(InputBox("Enter amount of copies"))

  if intAmtCopy <= 100 then
dblAmount = ( intAmtCopy * 0.05)
MsgBox("Total Cost is" & dblAmount)

  Elseif intAmtCopy > 100 then
dblAmount = ( intAmtCopy * 0.08)
MsgBox("Total Cost is" & dblAmount)
end if
Next
