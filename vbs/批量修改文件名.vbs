Function FilesTree(sPath)  
    Set oFso = CreateObject("Scripting.FileSystemObject")  
    Set oFolder = oFso.GetFolder(sPath)  
    Set oSubFolders = oFolder.SubFolders  
      
    Set oFiles = oFolder.Files  
    For Each oFile In oFiles  
        WScript.Echo oFile.Path
		oFso.MoveFile oFile.Path, replace(oFile.Path, "信息中心", "数据中心")
    Next  
      
    For Each oSubFolder In oSubFolders  
        FilesTree(oSubFolder.Path) 
    Next  
      
    Set oFolder = Nothing  
    Set oSubFolders = Nothing  
    Set oFso = Nothing  
End Function   


FilesTree("C:\Users\鑫\Desktop\VV文档") 