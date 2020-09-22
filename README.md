<div align="center">

## Remotely Run a SQL Server 7\.0 DTS Package from VB


</div>

### Description

Have you ever wondered how to remotely fire a DTS Package in SQL Server 7 from a Visual Basic Application? Me too... So through rigorous research and aggrevation I figured out a very simple way to do this.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Kesler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-kesler.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-kesler-remotely-run-a-sql-server-7-0-dts-package-from-vb__1-29503/archive/master.zip)





### Source Code

First: You create your DTS Package in SQL Server to do the job you need it to do.<br><br>
Second: You create a Stored Procedure similar to the one I have provided:<br><br>
<FONT SIZE=1><i>
CREATE PROC sp_SampleShell AS<BR>
EXEC master..xp_cmdshell 'C:\MSSQL7\BINN\DTSRun.exe /S [SERVERNAME] /N [DTSNAME] /E'
</i></FONT>
<BR><br>
this will execute the DTS Package via the xp_cmdshell provided by SQL. The DTSRun.exe will be found in your [MSSQL7\BINN] directory.<br><br>
Third: In your VB Program you create an ADO connection to your Database and use the following information in your program:<br><br>
<FONT SIZE=1><I>
'---------------------------<br>
IF RunPac(sp_SampleShell) = TRUE THEN<br>
  [do something]<br>
ELSE<br>
  [do something else]<br>
END IF<br>
'---------------------------<br>
Private Function RunPac(StProc As String) As Boolean
<br><br>
  Dim cnn As ADODB.Connection<br>
  Dim cmd As ADODB.Command<br><br>
  On Error GoTo Show_Err<br><br>
  Set cnn = New ADODB.Connection<br>
  Set cmd = New ADODB.Command<br><br>
  'set our connection constraints<br>
  With cnn<br>
    .ConnectionString = "DATA SOURCE=[DSN]"<br>
    .CursorLocation = adUseClient<br>
    .Open<br>
    'process the stored procedure command with no records to return<br>
    Set cmd = .Execute(StProc, , adExecuteNoRecords)<br>
  End With<br>
  cnn.Close<br>
  Set cnn = Nothing<br>
  Set cmd = Nothing<br>
  'if successful return true<br>
  RunPac = True<br>
  Exit Function<br>
Show_Err:<br>
  Debug.Print Err.Number & " - " & Err.Description<br>
  'if it fails return false<br>
  RunPac = False<br>
  cnn.Close<br>
  Set cnn = Nothing<br>
End Function<br>
</I></FONT>
<br><br>
<STRONG>And voila!!!</STRONG> You've just created a remote process for a DTS Package...<br><br><br>
I hope this helps someone else out as well. <br><br>
A very good point was made that there may be an easier way of doing this using the reference to DTS.dll. I tried using that method and had some issues with my environment so I needed to develop something that didn't care about the development environment. Also, this method is used more for those who not only develop their own VB Applications but also develop their own Stored Procedures as well.
<br><br>
I did not do another search in the past month or so regarding this so if this replicates anyone else I'm sorry, but this information did not exist when I originally needed it.

