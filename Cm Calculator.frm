Private Sub CommandButton34_Click()
    If (TextBox1.Value = vbNullString) Then
        TextBox1.Value = TextBox1.Value
    Else
        TextBox1.Text = Left(TextBox1.Text, Len(TextBox1.Text) - 1)
    End If
End Sub

Private Sub CommandButton12_Click()
TextBox1.Text = " "
End Sub

Private Sub CommandButton9_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Text = Log(TextBox1.Text)
End If
End Sub

Private Sub CommandButton32_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = Log(TextBox1.Text) / Log(10)
End If
End Sub

Private Sub CommandButton31_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = TextBox1.Value & "^"
End If
End Sub

Private Sub CommandButton30_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Text = Exp(TextBox1.Text)
End If
End Sub

Private Sub CommandButton8_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Text = Sqr(TextBox1.Text)
End If
End Sub

Private Sub CommandButton29_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Text = 1 / TextBox1.Text
End If
End Sub

Private Sub CommandButton27_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = TextBox1.Value & "/"
End If
End Sub

Private Sub CommandButton26_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = TextBox1.Value & "*"
End If
End Sub

Private Sub CommandButton25_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = TextBox1.Value & "-"
End If
End Sub

Private Sub CommandButton24_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = TextBox1.Value
Else
    TextBox1.Value = TextBox1.Value & "+"
End If
End Sub

Private Sub CommandButton23_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = "."
Else
    TextBox1.Value = TextBox1.Value & "."
End If
End Sub

Private Sub CommandButton20_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 1
Else
    TextBox1.Value = TextBox1.Value & 1
End If
End Sub

Private Sub CommandButton17_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 2
Else
    TextBox1.Value = TextBox1.Value & 2
End If
End Sub

Private Sub CommandButton21_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 3
Else
    TextBox1.Value = TextBox1.Value & 3
End If
End Sub

Private Sub CommandButton2_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 4
Else
    TextBox1.Value = TextBox1.Value & 4
End If
End Sub

Private Sub CommandButton19_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 5
Else
    TextBox1.Value = TextBox1.Value & 5
End If
End Sub

Private Sub CommandButton18_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 6
Else
    TextBox1.Value = TextBox1.Value & 6
End If
End Sub

Public Sub CommandButton7_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 7
Else
    TextBox1.Value = TextBox1.Value & 7
End If
End Sub

Private Sub CommandButton6_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 8
Else
    TextBox1.Value = TextBox1.Value & 8
End If
End Sub

Private Sub CommandButton5_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 9
Else
    TextBox1.Value = TextBox1.Value & 9
End If
End Sub

Private Sub CommandButton22_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 0
Else
    TextBox1.Value = TextBox1.Value & 0
End If
End Sub

Private Sub CommandButton3_Click()
If (TextBox1.Value = vbNullString) Then
    TextBox1.Value = 3.14159265
Else
    TextBox1.Value = TextBox1.Value & 3.14159265
End If
End Sub

Private Sub CommandButton35_Click()
If (TextBox1.Value = vbNullString) Then
        Dim Var As Object
        Var = MsgBox("Please enter an expression." & vbNewLine & "e.g., 5 + 4", vbExclamation, "Caution!")
Else
    TextBox1.Value = Application.Evaluate(TextBox1.Value)
End If
End Sub

Private Sub Label2_Click()
    Dim Var As Object
    Var = MsgBox("Cm Calculator." & vbNewLine & "Version 2.0.1.23 (Build 27)", 0, "About")
End Sub