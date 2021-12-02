Option Explicit
Dim productNum,quantitySold,totalRetail

DO WHILE productNum <> 6

productNum = CInt(InputBox("Enter product number or 6 to exit the program"&vbNewLine& _
"1 - Product 1, $2.98"&vbNewLine& _
"2 - Product 2, $4.50"&vbNewLine& _
"3 - Product 3, $9.98"&vbNewLine& _
"4 - Product 4, $4.49"&vbNewLine& _
"5 - Product 5, $6.87"))

if productNum = 6 then
exit do
end if

if productNum >=1 and productNum <=5 Then

quantitySold = CInt(InputBox("Enter the number of products sold"))
SELECT CASE productNum
CASE 1
totalRetail = quantitySold * 2.98
MsgBox("Total retail for products sold is: "&totalRetail)
CASE 2
totalRetail = quantitySold * 4.50
MsgBox("Total retail for products sold is: "&totalRetail)
CASE 3
totalRetail = quantitySold * 9.98
MsgBox("Total retail for products sold is: "&totalRetail)
CASE 4
totalRetail = quantitySold * 4.49
MsgBox("Total retail for products sold is: "&totalRetail)
CASE 5
totalRetail = quantitySold * 6.87
MsgBox("Total retail for products sold is: "&totalRetail)
END SELECT
else MsgBox("Invalid Input")
end if
LOOP
