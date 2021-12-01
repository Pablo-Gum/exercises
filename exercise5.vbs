'Declare Variables
Dim intCopy 'Input
Dim totalCopy 'Output

intCopy = cint(inputbox ("Enter Number of Copies "))

if intCopy > 0 and intCopy <= 100 then
totalCopy = intCopy * 0.05
MsgBox("Total price R"&totalCopy)

elseif intCopy > 100 then

totalCopy = 5 + (intCopy - 100) * 0.03
MsgBox("Total price R"&totalCopy)

else
MsgBox("Incorrect number of copies")

end If
