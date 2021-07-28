Attribute VB_Name = "Module1"
Sub ConnectSqlServer()
Attribute ConnectSqlServer.VB_ProcData.VB_Invoke_Func = "K\n14"

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim sql666 As String
    
                  
    sql666 = "SET DATEFORMAT dmy;" & _
"select ph.Name as 'Название фонограммы', ph_ref.val_str as 'Автор музыки', " & _
"ph_ref_2.val_str as 'Авто слов', COUNT('hist666.ph_id') as 'Кол-во сообщений в эфир', " & _
"ph_ref_3.val_str as 'Исполнитель (ФИО исполнителя или название коллектива)', ph_ref_4.val_str as 'Изготовитель фонограммы' from dbo.PH_PLAY_HISTORY hist666 " & _
"left join dbo.ph as ph " & _
"on hist666.ph_id = ph.id " & _
"left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Автор музыки') as ph_ref " & _
"on hist666.ph_id = ph_ref.ph_id " & _
"left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Автор слов') as ph_ref_2 " & _
"on hist666.ph_id = ph_ref_2.ph_id " & _
"left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Исполнитель') as ph_ref_3 " & _
"on hist666.ph_id = ph_ref_3.ph_id " & _
"left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Изготовитель фонограммы') as ph_ref_4 " & _
"on hist666.ph_id = ph_ref_4.ph_id " & _
"WHERE hist666.BlockId like 'PLAYLIST_RM%' and hist666.PlayTime between '01.07.2021' and '30.07.2021' and ph_ref.val_str is not null " & _
"group by hist666.PlayTime, ph.Name, hist666.ph_id, ph_ref.val_str, ph_ref_2.val_str, ph_ref_3.val_str, ph_ref_4.val_str " & _
"order by hist666.PlayTime"

    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=BAR-DCSRV-01\DS2SQLEXPRESS;" & _
                  "Initial Catalog=BAR-DS2;" & _
                  "Integrated Security=SSPI;"
    
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Open the connection and execute. Переменная sql - это и есть строка запроса к базе.
    conn.Open sConnString
    Set rs = conn.Execute(sql666)
    
  
    ' Check we have data.
    If Not rs.EOF Then
        ' Transfer result.
        Sheets(1).Range("A2").CopyFromRecordset rs
    ' Close the recordset
        rs.Close
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If

    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
End Sub

