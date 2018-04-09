Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	If KeyAscii < 48 Or KeyAscii > 57 Then
		KeyAscii = 0
	End If
End Sub