'Variable Declarations
Dim intMilesDriven
Dim intGallonsUsed
Dim intTotalMilesDriven
Dim intTotalGallonsUsed
Dim DblMilesPerGallon
Dim intNumTrips
Dim i

''user input
intNumTrips = CINT(INPUTBOX("Please enter the Number of trips."))

for i = 1 TO intNumTrips
intMilesDriven = CINT(INPUTBOX("Please enter the Miles driven for trip " & (i)))
intGallonsUsed = CINT(INPUTBOX("Please enter the Gallons Used for trip " & (i)))

''calc
DblMilesPerGallon = intMilesDriven/intGallonsUsed
MSGBOX("Miles per gallon for trip " & (i) & " is " & DblMilesPerGallon)
intTotalMilesDriven = intTotalMilesDriven + intMilesDriven
intTotalGallonsUsed = intTotalGallonsUsed + intGallonsUsed
Next

Dim DblTotalMilesPerGallon

DblTotalMilesPerGallon = intTotalMilesDriven/intTotalGallonsUsed
''Output to user
msgbox("The total mile per gallon for all trips is " & DblTotalMilesPerGallon)
