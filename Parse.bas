Attribute VB_Name = "Parse"
'This is a parse funcion, just like the split string
'i dunno why i use it, i prefer this i guess!
'VERY USEFULL!

Dim parse() As String
Sub ParseText(Text As String)
    parse = Split(Text, ":")
    
Form1.Text3 = parse(0)
Form1.Text4 = parse(1)
    

End Sub
