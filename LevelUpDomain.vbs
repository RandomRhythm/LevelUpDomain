'Takes a list of domains and output one unique domain structure for each unique second (or third) level domain
'Will also sort out IP addresses and invalid domains
'Output is deduplicated
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1

Dim dictTLD: set dictTLD = CreateObject("Scripting.Dictionary")
Dim dictSLD: set dictSLD = CreateObject("Scripting.Dictionary")
'http://data.iana.org/TLD/tlds-alpha-by-domain.txt
Dim dictPrev: set DictPrev = CreateObject("scripting.Dictionary")
Dim SecondLevelDict: Set SecondLevelDict = CreateObject("Scripting.Dictionary")
Dim ThirdLevelDict: Set ThirdLevelDict = CreateObject("Scripting.Dictionary")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim inputFile

inputFile = SelectFile()
msgbox "Select output directory"
strOutDir = fnShellBrowseForFolderVB
CurrentDirectory = GetFilePath(wscript.ScriptFullName)

AddTLDtoDict 'populate top level domain dict
AddSLDtoDict 'populate second level domain dict
LoadSecondDNS 'Load second level DNS
LoadThirdDNS 'Load third level DNS
if objFSO.fileexists(inputFile) then
  Set objFile = objFSO.OpenTextFile(inputFile)
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine
        strData = lcase(strData) 'force lowercase
        stroutDomain = ""
        intDomainDepth = 1 'grab top two domains
        if instr(strData, ".") > 0 and isIPaddress(strData) = False then 'has dot and is not IP address
          arrayLevelDomain = split(strData, ".")
          if dictTLD.exists("." & arrayLevelDomain(ubound(arrayLevelDomain))) then 'Country Code TLD
            if dictSLD.exists(arrayLevelDomain(ubound(arrayLevelDomain) -1)) then 'check second level domain
              intDomainDepth = intDomainDepth + 1 'grab top 3 domains
            end if
          end if
          for x = ubound(arrayLevelDomain) to (ubound(arrayLevelDomain) - intDomainDepth) step -1

            if stroutDomain = "" then
              stroutDomain = arrayLevelDomain(x)
            elseif ubound(arrayLevelDomain) > 2 and ThirdLevelDict.exists(arrayLevelDomain(x - 1) & "." & arrayLevelDomain(x) & "." & stroutDomain) then 'known third level domain
              stroutDomain = arrayLevelDomain(x - 2) & "." & arrayLevelDomain(x - 1) & "."  & arrayLevelDomain(x) & "." & stroutDomain
              msgbox "four level: " & stroutDomain
              exit for 'confirmed 4 level domain           
            elseif ubound(arrayLevelDomain) > 1 and SecondLevelDict.exists(arrayLevelDomain(x) & "." & stroutDomain) then 'known second level domain
              stroutDomain = arrayLevelDomain(x - 1) & "." & arrayLevelDomain(x) & "." & stroutDomain
              exit for 'confirmed 3 level domain
              msgbox "third level: " & stroutDomain
            else
              stroutDomain = arrayLevelDomain(x) & "." & stroutDomain
            end if
          next
			
				
            if dictPrev.exists(stroutDomain) = False then
              dictPrev.add stroutDomain, 0
              logdata strOutDir & "\LevelUP_Domains.txt", stroutDomain, False
              logdata strOutDir & "\Domain_Sample.txt", strData, False
            else 'prevalence if SLD
              dictPrev.item(stroutDomain) = dictPrev.item(stroutDomain) + 1
            end if
        else 'not domain 
          if isIPaddress(strData) = False then   
            logdata strOutDir & "\Invalid_Domain_IP.txt", strData, False
          else
            logdata strOutDir & "\IP_Addresses.txt", strData, False
          end if
        end if

        on error goto 0
    end if
  loop
end if

for each domain in dictPrev
  logdata strOutDir & "\DomainPrev.csv", domain & "," & dictPrev.item(domain), false

next

msgbox "Done"

function LogData(TextFileName, TextToWrite,EchoOn)

Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if TextFileName = "" then
  msgbox "No file path passed to LogData"
  exit function
end if
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      on error resume next
      fsoLogData.CreateTextFile TextFileName, True
      if err.number <> 0 and err.number <> 53 then 
        logdata CurrentDirectory & "\VT_Error.log", Date & " " & Time & " Error logging to " & TextFileName & " - " & err.description,False 
        objShellComplete.popup err.number & " " & err.description & vbcrlf & TextFileName,,"Logging error", 30
        exit function
      end if
      on error goto 0
  End If
if TextFileName <> "" then

  on error resume next
  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  WriteTextFile.WriteLine TextToWrite
  WriteTextFile.Close
  if err.number <> 0 then 
    on error goto 0
    
  Dim objStream
  Set objStream = CreateObject("ADODB.Stream")
  objStream.CharSet = "utf-16"
  objStream.Open
  objStream.WriteText TextToWrite
  on error resume next
  objStream.SaveToFile TextFileName, 2
  if err.number <> 0 then msgbox err.number & " - " & err.message & " Problem writing to " & TextFileName
  if err.number <> 0 then 
    objShellComplete.popup "problem writting text: " & TextToWrite, 30
    logdata CurrentDirectory & "\VT_Error.log", Date & " " & Time & " problem writting text: " & TextToWrite,False 
  end if
  on error goto 0
  Set objStream = nothing
  end if
end if
Set fsoLogData = Nothing
End Function




Function GetFilePath (ByVal FilePathName)
found = False

Z = 1

Do While found = False and Z < Len((FilePathName))

 Z = Z + 1

         If InStr(Right((FilePathName), Z), "\") <> 0 And found = False Then
          mytempdata = Left(FilePathName, Len(FilePathName) - Z)
          
             GetFilePath = mytempdata

             found = True

        End If      

Loop

end Function


Function isIPaddress(strIPaddress)
DIm arrayTmpquad
Dim boolReturn_isIP
boolReturn_isIP = True
if instr(strIPaddress,".") then
  arrayTmpquad = split(strIPaddress,".")
  for each item in arrayTmpquad
    if isnumeric(item) = false then boolReturn_isIP = false
  next
else
  boolReturn_isIP = false
end if
if boolReturn_isIP = false then
	boolReturn_isIP = isIpv6(strIPaddress)
end if
isIPaddress = boolReturn_isIP
END FUNCTION

Function IsIPv6(TestString)

    Dim sTemp
    Dim iLen
    Dim iCtr
    Dim sChar
    
    if instr(TestString, ":") = 0 then 
		IsIPv6 = false
		exit function
	end if
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            if isnumeric(sChar) or "a"= lcase(sChar) or "b"= lcase(sChar) or "c"= lcase(sChar) or "d"= lcase(sChar) or "e"= lcase(sChar) or "f"= lcase(sChar) or ":" = sChar then
              'allowed characters for hash (hex)
            else
              IsIPv6 = False
              exit function
            end if
        Next
    
    IsIPv6 = True
    else
      IsIPv6 = False
    End If
    
End Function


sub AddSLDtoDict
dictSLD.add "co",0
dictSLD.add "com",0
dictSLD.add "net",0
dictSLD.add "org",0
dictSLD.add "edu",0
dictSLD.add "gov",0
dictSLD.add "asn",0
dictSLD.add "id",0
dictSLD.add "csiro",0

end sub

sub AddTLDtoDict

dictTLD.add ".af",0
dictTLD.add ".ax",0
dictTLD.add ".al",0
dictTLD.add ".dz",0
dictTLD.add ".as",0
dictTLD.add ".ad",0
dictTLD.add ".ao",0
dictTLD.add ".ai",0
dictTLD.add ".aq",0
dictTLD.add ".ag",0
dictTLD.add ".ar",0
dictTLD.add ".am",0
dictTLD.add ".aw",0
dictTLD.add ".ac",0
dictTLD.add ".au",0
dictTLD.add ".at",0
dictTLD.add ".az",0
dictTLD.add ".bs",0
dictTLD.add ".bh",0
dictTLD.add ".bd",0
dictTLD.add ".bb",0
dictTLD.add ".eus",0
dictTLD.add ".by",0
dictTLD.add ".be",0
dictTLD.add ".bz",0
dictTLD.add ".bj",0
dictTLD.add ".bm",0
dictTLD.add ".bt",0
dictTLD.add ".bo",0
dictTLD.add ".bq",0
dictTLD.add ".ba",0
dictTLD.add ".bw",0
dictTLD.add ".bv",0
dictTLD.add ".br",0
dictTLD.add ".io",0
dictTLD.add ".vg",0
dictTLD.add ".bn",0
dictTLD.add ".bg",0
dictTLD.add ".bf",0
dictTLD.add ".mm",0
dictTLD.add ".bi",0
dictTLD.add ".kh",0
dictTLD.add ".cm",0
dictTLD.add ".ca",0
dictTLD.add ".cv",0
dictTLD.add ".cat",0
dictTLD.add ".ky",0
dictTLD.add ".cf",0
dictTLD.add ".td",0
dictTLD.add ".cl",0
dictTLD.add ".cn",0
dictTLD.add ".cx",0
dictTLD.add ".cc",0
dictTLD.add ".co",0
dictTLD.add ".km",0
dictTLD.add ".cd",0
dictTLD.add ".cg",0
dictTLD.add ".ck",0
dictTLD.add ".cr",0
dictTLD.add ".ci",0
dictTLD.add ".hr",0
dictTLD.add ".cu",0
dictTLD.add ".cw",0
dictTLD.add ".cy",0
dictTLD.add ".cz",0
dictTLD.add ".dk",0
dictTLD.add ".dj",0
dictTLD.add ".dm",0
dictTLD.add ".do",0
dictTLD.add ".tl",0
dictTLD.add ".ec",0
dictTLD.add ".eg",0
dictTLD.add ".sv",0
dictTLD.add ".gq",0
dictTLD.add ".er",0
dictTLD.add ".ee",0
dictTLD.add ".et",0
dictTLD.add ".eu",0
dictTLD.add ".fk",0
dictTLD.add ".fo",0
dictTLD.add ".fm",0
dictTLD.add ".fj",0
dictTLD.add ".fi",0
dictTLD.add ".fr",0
dictTLD.add ".gf",0
dictTLD.add ".pf",0
dictTLD.add ".tf",0
dictTLD.add ".ga",0
dictTLD.add ".gal",0
dictTLD.add ".gm",0
dictTLD.add ".ps",0
dictTLD.add ".ge",0
dictTLD.add ".de",0
dictTLD.add ".gh",0
dictTLD.add ".gi",0
dictTLD.add ".gr",0
dictTLD.add ".gl",0
dictTLD.add ".gd",0
dictTLD.add ".gp",0
dictTLD.add ".gu",0
dictTLD.add ".gt",0
dictTLD.add ".gg",0
dictTLD.add ".gn",0
dictTLD.add ".gw",0
dictTLD.add ".gy",0
dictTLD.add ".ht",0
dictTLD.add ".hm",0
dictTLD.add ".hn",0
dictTLD.add ".hk",0
dictTLD.add ".hu",0
dictTLD.add ".is",0
dictTLD.add ".in",0
dictTLD.add ".id",0
dictTLD.add ".ir",0
dictTLD.add ".iq",0
dictTLD.add ".ie",0
dictTLD.add ".im",0
dictTLD.add ".il",0
dictTLD.add ".it",0
dictTLD.add ".jm",0
dictTLD.add ".jp",0
dictTLD.add ".je",0
dictTLD.add ".jo",0
dictTLD.add ".kz",0
dictTLD.add ".ke",0
dictTLD.add ".ki",0
dictTLD.add "not",0
dictTLD.add ".kw",0
dictTLD.add ".kg",0
dictTLD.add ".la",0
dictTLD.add ".lv",0
dictTLD.add ".lb",0
dictTLD.add ".ls",0
dictTLD.add ".lr",0
dictTLD.add ".ly",0
dictTLD.add ".li",0
dictTLD.add ".lt",0
dictTLD.add ".lu",0
dictTLD.add ".mo",0
dictTLD.add ".mk",0
dictTLD.add ".mg",0
dictTLD.add ".mw",0
dictTLD.add ".my",0
dictTLD.add ".mv",0
dictTLD.add ".ml",0
dictTLD.add ".mt",0
dictTLD.add ".mh",0
dictTLD.add ".mq",0
dictTLD.add ".mr",0
dictTLD.add ".mu",0
dictTLD.add ".yt",0
dictTLD.add ".mx",0
dictTLD.add ".md",0
dictTLD.add ".mc",0
dictTLD.add ".mn",0
dictTLD.add ".me",0
dictTLD.add ".ms",0
dictTLD.add ".ma",0
dictTLD.add ".mz",0
dictTLD.add ".na",0
dictTLD.add ".nr",0
dictTLD.add ".np",0
dictTLD.add ".nl",0
dictTLD.add ".nc",0
dictTLD.add ".nz",0
dictTLD.add ".ni",0
dictTLD.add ".ne",0
dictTLD.add ".ng",0
dictTLD.add ".nu",0
dictTLD.add ".nf",0
dictTLD.add ".nc.tr",0
dictTLD.add ".kp",0
dictTLD.add ".mp",0
dictTLD.add ".no",0
dictTLD.add ".om",0
dictTLD.add ".pk",0
dictTLD.add ".pw",0
dictTLD.add ".pa",0
dictTLD.add ".pg",0
dictTLD.add ".py",0
dictTLD.add ".pe",0
dictTLD.add ".ph",0
dictTLD.add ".pn",0
dictTLD.add ".pl",0
dictTLD.add ".pt",0
dictTLD.add ".pr",0
dictTLD.add ".qa",0
dictTLD.add ".ro",0
dictTLD.add ".ru",0
dictTLD.add ".rw",0
dictTLD.add ".re",0
dictTLD.add ".bl",0
dictTLD.add ".sh",0
dictTLD.add ".kn",0
dictTLD.add ".lc",0
dictTLD.add ".mf",0
dictTLD.add ".pm",0
dictTLD.add ".vc",0
dictTLD.add ".ws",0
dictTLD.add ".sm",0
dictTLD.add ".st",0
dictTLD.add ".sa",0
dictTLD.add ".sn",0
dictTLD.add ".rs",0
dictTLD.add ".sc",0
dictTLD.add ".sl",0
dictTLD.add ".sg",0
dictTLD.add ".sx",0
dictTLD.add ".sk",0
dictTLD.add ".si",0
dictTLD.add ".sb",0
dictTLD.add ".so",0
dictTLD.add ".za",0
dictTLD.add ".gs",0
dictTLD.add ".kr",0
dictTLD.add ".ss",0
dictTLD.add ".es",0
dictTLD.add ".lk",0
dictTLD.add ".sd",0
dictTLD.add ".sr",0
dictTLD.add ".sj",0
dictTLD.add ".sz",0
dictTLD.add ".se",0
dictTLD.add ".ch",0
dictTLD.add ".sy",0
dictTLD.add ".tw",0
dictTLD.add ".tj",0
dictTLD.add ".tz",0
dictTLD.add ".th",0
dictTLD.add ".tg",0
dictTLD.add ".tk",0
dictTLD.add ".to",0
dictTLD.add ".tt",0
dictTLD.add ".tn",0
dictTLD.add ".tr",0
dictTLD.add ".tm",0
dictTLD.add ".tc",0
dictTLD.add ".tv",0
dictTLD.add ".ug",0
dictTLD.add ".ua",0
dictTLD.add ".ae",0
dictTLD.add ".uk",0
dictTLD.add ".us",0
dictTLD.add ".vi",0
dictTLD.add ".uy",0
dictTLD.add ".uz",0
dictTLD.add ".vu",0
dictTLD.add ".va",0
dictTLD.add ".ve",0
dictTLD.add ".vn",0
dictTLD.add ".wf",0
dictTLD.add ".eh",0
dictTLD.add ".ye",0
dictTLD.add ".zm",0
dictTLD.add ".zw",0
end sub

Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&ælig;-4ba3-bca5-ec349df65ef6

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function

function fnShellBrowseForFolderVB()
    dim objShell
    dim ssfWINDOWS
    dim objFolder
    
    ssfWINDOWS = 36
    set objShell = CreateObject("shell.application")
        set objFolder = objShell.BrowseForFolder(0, "Example", 0, ssfDRIVES)
            if (not objFolder is nothing) then
               set oFolderItem = objFolder.items.item
               fnShellBrowseForFolderVB = oFolderItem.Path 
            end if
        set objFolder = nothing
    set objShell = nothing
end function



Sub LoadSecondDNS()'load list from http://george.surbl.org/two-level-tlds
if objFSO.fileexists(CurrentDirectory & "\two-level-tlds.txt") then
  Set objFile = objFSO.OpenTextFile(CurrentDirectory & "\two-level-tlds.txt")
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine 
        on error goto 0
          SecondLevelDict.add strData, 1
    end if
  loop
end if
end sub


Sub LoadThirdDNS() 'loads list from http://www.surbl.org/static/three-level-tlds
if objFSO.fileexists(CurrentDirectory & "\three-level-tlds.txt") then
  Set objFile = objFSO.OpenTextFile(CurrentDirectory & "\three-level-tlds.txt")
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine 
        on error goto 0
          ThirdLevelDict.add strData, 1
    end if
  loop
end if
end sub