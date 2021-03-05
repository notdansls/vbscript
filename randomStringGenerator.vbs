option explicit

' Integer declarations
Dim intCurrentLength, intMin, intMax, intIntermediate, intRand , intCapatalise, intResult, rand
' Array declaration
Dim arrAlphabet
arrAlphabet = Array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z")

' String declarations
Dim strCase, strFinalString, strGeneratedWord, strCopy
' Object declaration
Dim objIE

function generateNumber(intMin,intMax)
	Randomize()
	generateNumber = Int((intMax-intMin+1)*Rnd+intMin)
end function

function decideCapitalOrLower
	if generateNumber (0,1) = 0 then
		decideCapitalOrLower = "lowerCase"
	else
		decideCapitalOrLower = "upperCase"
	end if
end function

function generateLetter
	intMin = 0
	intMax = 25
	Randomize()
	generateLetter = arrAlphabet(Int((intMax-intMin+1)*Rnd+intMin))
end function

function generateString(intTargetLength)
	intCurrentLength = 0
	strFinalString = ""
	
	do while intCurrentLength < intTargetLength
		strCase = decideCapitalOrLower
		
		if strCase = "lowerCase" then
			strFinalString = strFinalString + Lcase(generateLetter)
			
		elseif strCase = "upperCase" then
			strFinalString = strFinalString + Ucase(generateLetter)
		else
			msgbox "An error has occurred "
		end if
		intCurrentLength = intCurrentLength + 1
	loop
	
	generateString = strFinalString
end function

wscript.echo generateString (8)