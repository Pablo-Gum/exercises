' declare variables
Dim num_bagels' input
Dim total_cost ' output



num_bagels = CInt(InputBox("Please enter number of bagels"))
if num_bagels > 0 and num_bagels < 6 then
total_cost = num_bagels * 0.75
MsgBox("Total cost is " & total_cost)
end if
if num_bagels >= 6 Then
total_cost = num_bagels * 0.60
MsgBox("Total cost is " & total_cost)
end if



if num_bagels < 1 Then
MsgBox("Invalid entry")
end if
