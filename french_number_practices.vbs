Public Function GetRandNumber(min, max)
	Randomize
	GetRandNumber = Int((max-min+1)*Rnd+min)
End Function
Dim number_str, sapi, practice_number, scale, i, answer
i=0
practice_number=CInt(InputBox("Input number of exercises: ", "Input"))
scale = CInt(InputBox("Input number range: from 0 to ","Input"))
Set sapi=CreateObject("sapi.spvoice")
Set sapi.Voice=sapi.GetVoices.Item(2)

Do Until i=practice_number

number_str=CStr(GetRandNumber(0, scale))
sapi.Speak number_str
answer=CStr(InputBox("Input your answer: ", "Ans"))
If answer=number_str Then MsgBox "Correct" Else MsgBox "Wrong!"&" It's "&number_str

i=i+1
Loop
