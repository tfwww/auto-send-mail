
xlsSrc := "E:\heyalan\files-Gold Seagull\Jennifer\record.xlsx"
adOpenStatic := 3
adLockOptimistic := 3
adCmdText := 0x0001

objConnection := ComObjCreate("ADODB.Connection")
objRecordSet := ComObjCreate("ADODB.Recordset")

objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" . xlsSrc . "; Extended Properties='Excel 12.0;HDR=Yes'")

objRecordset.Open("Select * FROM [custormers$]", objConnection, adOpenStatic, adLockOptimistic, adCmdText)

SendMail(address, subject, name) {
    body :=  "Dear " name  ",<br>This is Jennifer from FUHUIDA TECHNOLOGY, we are a PCB manufacture in China.<br>We focus on high-tech enterprise, especially in 1-30 layers PCB prototype, and batch production with competitive price.<br>We can provide high quality and professional service to you. If you have any question or quotation please contact me freely. I believe we can have a good beginning of cooperation.<br>Thanks for your valuable time and awaiting for your reply.<br>"

    Run, mailto:%address%?subject=%subject%&body=%body%
    Sleep 1000
    Send {Ctrl down}
    Sleep 500
    Send {Enter}
    Send {Ctrl up}
    Sleep 10000
}

; SendMail(address, subject, name)

while !objRecordset.EOF
{
    address := objRecordset.Fields.Item("Mail").Value
    name := objRecordset.Fields.Item("Contact").Value
    subject := "Quotation"
    Sleep 5000
    SendMail(address, subject, name)
    objRecordset.MoveNext
}
