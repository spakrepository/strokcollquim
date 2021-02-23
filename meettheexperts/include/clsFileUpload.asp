<%
Class clsFileUpload
'=================================================================================================
'This class will parse the binary contents of the request, and populate the Form Field collection.
'=================================================================================================
	Private objForm					''(R)	Collection of all the form fields including file field
	Private objFiles				''()	Collection of all the file fields only
	Private lngMaxSize				''(RW)	maximum file size accepted
	Private strFileType				''(RW)	specific file type accepted
	Private strImageDetailSupport	''(R)	list of all the image files support for width and height(THIS IS INFORMATIVE ONLY)
	Private WebImage				''(RW)	Boolean --All the strImageDetailSupport
	
	Private Sub Class_Initialize()
		Set objFiles	= New clsCollection
		Set objForm		= New clsCollection
		lngMaxSize		= 0
		strFileType		= "All"
		strImageDetailSupport	= "GIF;JPG;JPEG;PNG;WMF;BMP;"
		WebImage = False
		CALL ParseRequest()
	End Sub

	''PUBLIC PROPERTY
	Public Property Get Form() : Set Form = objForm :	End Property
	Public Property Get MaxFileSize() :	MaxFileSize = lngMaxSize :	End Property
	Public Property Let MaxFileSize(v) : lngMaxSize = v :	End Property
	Public Property Get FileType() : FileType = strFileType : End Property
	Public Property Let	FileType(v) : strFileType = v : End Property
	Public Property Get WebImages() : WebImages = WebImage : End Property
	Public Property Let	WebImages(v) : WebImage = v : End Property
	''PRIVATE PROPERTY
	Private Property Get Files() :	Set Files = objFiles : End Property
	
	Private Function CheckImage(strFileType)
		CheckImage=False
		If (WebImage) Then
			CheckImage=True
		ElseIf (strComp(strFileType,"JPEG",1)=0) Then
			CheckImage=True 
		ElseIF (strComp(strFileType,"JPG",1)=0) Then
			CheckImage=True
		ElseIf (strComp(strFileType,"WMF",1)=0) Then
			CheckImage=True
		ElseIf (strComp(strFileType,"GIF",1)=0) Then
			CheckImage=True
		ElseIf (strComp(strFileType,"BMP",1)=0) Then
			CheckImage=True
		ElseIf (strComp(strFileType,"PNG",1)=0) Then
			CheckImage=True
		End If
	End Function

	Private Function BStr2UStr(BStr)
	'Byte string to Unicode string conversion
		Dim lngLoop
		BStr2UStr = ""
		For lngLoop = 1 to LenB(BStr)
			BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr,lngLoop,1))) 
		Next
	End Function
	
	Private Function UStr2Bstr(UStr)
	'Unicode string to Byte string conversion
		Dim lngLoop
		Dim strChar
		UStr2Bstr = ""
		For lngLoop = 1 to Len(UStr)
			strChar = Mid(UStr, lngLoop, 1)
			UStr2Bstr = UStr2Bstr & ChrB(AscB(strChar))
		Next
	End Function
	
	Private Function URLDecode(Expression)
	'Why doesn't ASP provide this functionality for us?
		Dim strSource, strTemp, strResult
		Dim lngPos
		strSource = Replace(Expression, "+", " ")
		For lngPos = 1 To Len(strSource)
			strTemp = Mid(strSource, lngPos, 1)
			If strTemp = "%" Then
				If lngPos + 2 < Len(strSource) Then
					strResult = strResult & Chr(CInt("&H" & Mid(strSource, lngPos + 1, 2)))
					lngPos = lngPos + 2
				End If
			Else
				strResult = strResult & strTemp
			End If
		Next
		URLDecode = strResult
	End Function	

	Private Sub AddFormField(strName,strValue)
		If objForm.Add(strName,strValue,False)=False then
			atValue = objForm.Item(strName)
			objForm.Add strName,( atValue &", "& strValue ),True 
		Else
			objForm.Add strName,strValue,True
		End if
	End Sub
	
	Private Sub ParseRequest()
		Dim lngTotalBytes, lngPosBeg, lngPosEnd, lngPosBoundary, lngPosTmp, lngPosFileName
		Dim strBRequest, strBBoundary, strBContent
		Dim strName, strFileName, strContentType, strValue, strTemp
		Dim objFile
				
		'Grab the entire contents of the Request as a Byte string
		lngTotalBytes = Request.TotalBytes
		strBRequest = Request.BinaryRead(lngTotalBytes)
		
		'Find the first Boundary
		lngPosBeg = 1
		lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2Bstr(Chr(13)))
		If lngPosEnd > 0 Then
			strBBoundary = MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg)
			lngPosBoundary = InStrB(1, strBRequest, strBBoundary)
		End If
		If strBBoundary = "" Then
		'The form must have been submitted *without* ENCTYPE="multipart/form-data"
		'But since we already called Request.BinaryRead, we can no longer access
		'the Request.Form collection, so we need to parse the request and populate
		'our own form collection.
			lngPosBeg = 1
			lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr("&"))
			Do While lngPosBeg < LenB(strBRequest)
				'Parse the element and add it to the collection
				strTemp = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
				lngPosTmp = InStr(1, strTemp, "=")
				strName = URLDecode(Left(strTemp, lngPosTmp - 1))
				strValue = URLDecode(Right(strTemp, Len(strTemp) - lngPosTmp))
				''Form Field Elements is added
				AddFormField strName,strValue
				'Find the next element
				lngPosBeg = lngPosEnd + 1
				lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr("&"))
				If lngPosEnd = 0 Then lngPosEnd = LenB(strBRequest) + 1
			Loop
		Else
		'The form was submitted with ENCTYPE="multipart/form-data"
		'Loop through all the boundaries, and parse them into either the
		'Form or Files collections.
			Do Until (lngPosBoundary = InStrB(strBRequest, strBBoundary & UStr2Bstr("--")))
				'Get the element name
				lngPosTmp = InStrB(lngPosBoundary, strBRequest, UStr2BStr("Content-Disposition"))
				lngPosTmp = InStrB(lngPosTmp, strBRequest, UStr2BStr("name="))
				lngPosBeg = lngPosTmp + 6
				lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr(Chr(34)))
				strName = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
				'Look for an element named 'filename'
				lngPosFileName = InStrB(lngPosBoundary, strBRequest, UStr2BStr("filename="))
				'If found, we have a file, otherwise it is a normal form element
				If lngPosFileName <> 0 And lngPosFileName < InStrB(lngPosEnd, strBRequest, strBBoundary) Then 'It is a file
					'Get the FileName
					lngPosBeg = lngPosFileName + 10
					lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr(chr(34)))
					strFileName = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Get the ContentType
					lngPosTmp = InStrB(lngPosEnd, strBRequest, UStr2BStr("Content-Type:"))
					lngPosBeg = lngPosTmp + 14
					lngPosEnd = InstrB(lngPosBeg, strBRequest, UStr2BStr(chr(13)))
					strContentType = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Get the Content
					lngPosBeg = lngPosEnd + 4
					lngPosEnd = InStrB(lngPosBeg, strBRequest, strBBoundary) - 2
					strBContent = MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg)
					If strFileName <> "" And strBContent <> "" Then
						'Create the File object, and add it to the Files collection
						Set objFile = New clsFile
						objFile.Name = strName
						objFile.FileName = Right(strFileName, Len(strFileName) - InStrRev(strFileName, "\"))
						strFileType = UCase(Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".")))
						objFile.ContentType = strContentType
						objFile.Blob = strBContent
						If CheckImage(strFileType) then	
							If strFileType="JPG" Or strFileType="JPEG" then
								strFileTemp = BStr2UStr(strBContent)
								getDimension strFileTemp, Height_, Width_
								'Response.Write "Hello"
								'getDimensionJPG strBContent,Height_,Width_
							Else
								getDimension BStr2UStr(Left(strBContent,200)),Height_,Width_
							End if
							objFile.Height = Height_
							objFile.Width = Width_
						End if
						objForm.Add strName, objFile, True
					End If
				Else 'It is a form element
					'Get the value of the form element
					lngPosTmp = InStrB(lngPosTmp, strBRequest, UStr2BStr(chr(13)))
					lngPosBeg = lngPosTmp + 4
					lngPosEnd = InStrB(lngPosBeg, strBRequest, strBBoundary) - 2
					strValue = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Add the element to the collection
					''objForm.Add strName, strValue
					AddFormField strName,strValue
				End If
				'Move to Next Element
				lngPosBoundary = InStrB(lngPosBoundary + LenB(strBBoundary), strBRequest, strBBoundary)
			Loop
		End If
	End Sub
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':::  Functions to convert two bytes to a numeric value (long)   :::
	':::  (both little-endian and big-endian)                        :::
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	function lngConvert(strTemp)
		lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
	end function
	function lngConvert2(strTemp)
		lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
	end function
	function lngConvert2J(strTemp)
		lngConvert2J = clng(ascB(rightB(strTemp, 1)) + ((ascB(leftB(strTemp, 1)) * 256)))
	end function

	function getDimension(ByRef strFile,Height,Width)
	   dim strPNG,strGIF,strBMP,strType
	   strType = ""
	   strImageType = "(unknown)"
	   strPNG = chr(137) & chr(80) & chr(78)
	   strGIF = "GIF"
	   strBMP = chr(66) & chr(77)
	   strType = Left(strFile,3)
	   'Response.Write strType
	   if strType = strGIF then				' is GIF
	      strImageType = "GIF"
	      Width = lngConvert(Mid(strFile, 7, 2))
	      Height = lngConvert(Mid(strFile, 9, 2))
	   elseif left(strType, 2) = strBMP then		' is BMP
	      strImageType = "BMP"
	      Width = lngConvert(Mid(strFile, 19, 2))
	      Height = lngConvert(Mid(strFile, 23, 2))
	   elseif strType = strPNG then			' Is PNG
	      strImageType = "PNG"
	      Width = lngConvert2(Mid(strFile, 19, 2))
	      Height = lngConvert2(Mid(strFile, 23, 2))
	   else
			lngSize = len(strFile)
			flgFound = 0
			strTarget = chr(255) & chr(216) & chr(255)
			flgFound = instr(strFile, strTarget)
			if flgFound = 0 then
			   exit function
			end if
			strImageType = "JPG"
			lngPos = flgFound + 2
			ExitLoop = false
			do while ExitLoop = False and lngPos < lngSize
				ascTemp = asc(mid(strFile, lngPos, 1))
				do while ascTemp = 255 and lngPos < lngSize
					lngPos = lngPos + 1
					ascTemp = asc(mid(strFile, lngPos, 1))
				loop
				if ascTemp < 192 or ascTemp > 195 then
					lngMarkerSize = lngConvert2(mid(strFile, lngPos + 1, 2))
					''Response.Write "Marker Size = " & lngMarkerSize & "<BR>Position = " & lngPos &"<BR>"
					lngPos = lngPos + lngMarkerSize  + 1
				else
					'Response.Write lngPos &"<BR>"
					'Response.Write ascTemp &"<BR>" 
					ExitLoop = True
				end if
			loop
			if ExitLoop = False then
			   Width = -1
			   Height = -1
			else
				'lngPos = InStr(1,strFile,Chr(194),0)
				Height = lngConvert2(mid(strFile, lngPos + 4, 2))
				Width = lngConvert2(mid(strFile, lngPos + 6, 2))
			end if
	   end if
	end function
	Function getDimensionJPG(strFile,Height,Width)
	      lngSize = LenB(strFile)
	      flgFound = 0
	      strTarget = (chr(255)) & (chr(216)) & (chr(255))
	      flgFound = InStrB(1,strFile, strTarget,0)
	      if flgFound = 0 then
	         Response.Write "Not Found = " & strTarget
	         exit function
	      end if
	      strImageType = "JPG"
	      lngPos = flgFound + 2
	      ExitLoop = False
	      Do while (ExitLoop = False and lngPos < lngSize)
	         ascTemp = ascw(midB(strFile, lngPos, 1))
	         Do while (ascTemp = 255 and lngPos < lngSize)
				ascTemp = ascw(midB(strFile, lngPos, 1))
	            lngPos = lngPos + 1
	         Loop
	         if ascTemp < 192 or ascTemp > 195 then
	            lngMarkerSize = lngConvert2J(midB(strFile, lngPos + 1, 2))
	            lngPos = lngPos + lngMarkerSize  + 1
	         else
	            ExitLoop = True
	         end if
	     loop
	     if ExitLoop = False then
	        Width = -1
	        Height = -1
	     else
	        Height = lngConvert2J(midB(strFile, lngPos + 4, 2))
	        Width = lngConvert2J(midB(strFile, lngPos + 6, 2))
	     end if
	End Function
End Class


Class clsCollection
	Private objDicItems

	Private Sub Class_Initialize()
		Set objDicItems = Server.CreateObject("Scripting.Dictionary")
		objDicItems.CompareMode = vbTextCompare
	End Sub
	
	Public Property Get Count() : Count = objDicItems.Count :	End Property
	
	Public Default Function Item(Index)
		Dim arrItems
		If IsNumeric(Index) Then
			arrItems = objDicItems.Items
			If IsObject(arrItems(Index)) Then
				Set Item = arrItems(Index)
			Else
				Item = arrItems(Index)
			End If
		Else
			If objDicItems.Exists(Index) Then
				If IsObject(objDicItems.Item(Index)) Then
					Set Item = objDicItems.Item(Index)
				Else
					Item = objDicItems.Item(Index)
				End If
			else
				Item=""
			End If
		End If
	End Function
	
	Public Function Key(Index)
		Dim arrKeys
		If IsNumeric(Index) Then
			arrKeys = objDicItems.Keys
			Key = arrKeys(Index)
		End If
	End Function

	Public Function Add(Name, Value, overWrite)
		If objDicItems.Exists(Name) Then
			if overWrite then
				objDicItems.Item(Name) = Value
				Add = True
			Else
				Add = False
			End if
		Else
			objDicItems.Add Name, Value
			Add = True
		End If
	End Function
End Class

Class clsFile
	Private strName
	Private strContentType
	Private strFileName
	Private bBlob
	Public isBig
	Private height_
	Private width_
	Private Size_
	
	Private Sub Class_Initialize()
		isBig = False
		height_ = 0
		width_ = 0
	End Sub

	Public Property Get Name() : Name = strName : End Property
	Public Property Let Name(vIn) : strName = vIn : End Property
	Public Property Get ContentType() : ContentType = strContentType : End Property
	Public Property Let ContentType(vIn) : strContentType = vIn : End Property
	Public Property Get FileName() : FileName = strFileName : End Property
	Public Property Let FileName(vIn) : strFileName = vIn : End Property
	Public Property Get Blob() : Blob = bBlob : End Property
	
	Public Property Let Blob(vIn) 
		bBlob = vIn 
		Size_ = LenB(vIn)
	End Property
	
	Public Property Get Width() : Width = width_ : End Property
	Public Property Let Width(vIn) : width_ = vIn : End Property
	Public Property Get Height() : Height = height_ : End Property
	Public Property Let Height(vIn) : height_ = vIn : End Property		
	Public Property Get Size() : Size = size_ : End Property

	Public Function Save(Path)
		Dim objFSO, objFSOFile,lngLoop
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		Set objFSOFile = objFSO.CreateTextFile(objFSO.BuildPath(Path, strFileName))
		If Err.number<>0 then 
			Set objFSO = Nothing
			Save=Err.Description
			Exit Function
		End if
		For lngLoop = 1 to LenB(bBlob)
			objFSOFile.Write Chr(AscB(MidB(bBlob, lngLoop, 1)))
		Next
		objFSOFile.Close	
		Save=True
	End Function
	Public Function SaveInVirtual(Path)
		Dim objFSO, objFSOFile,lngLoop
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		Set objFSOFile = objFSO.CreateTextFile(objFSO.BuildPath(Server.MapPath(Path), strFileName))
		If Err.number<>0 then 
			Set objFSO = Nothing
			SaveInVirtual=Err.Description
			Exit Function
		End if
		For lngLoop = 1 to LenB(bBlob)
			objFSOFile.Write Chr(AscB(MidB(bBlob, lngLoop, 1)))
		Next
		objFSOFile.Close	
	End Function
	Public Function GetBlob()
		Dim D
		For lngLoop = 1 to LenB(bBlob)
			D = D & Chr(AscB(MidB(bBlob, lngLoop, 1)))
		Next
		GetBlob = D
	End Function

End Class


Class clsImage
	Private Height_
	Private Width_
	Private Colors_
	Private ImageType_
	Private Sub Class_Initialize()
		height_ = -1
		width_ = -1
		Colors_ = -1
		ImageType_ = ""
	End Sub

	Public Property Get Width() : Width = width_ : End Property
	Public Property Let Width(vIn) : width_ = vIn : End Property
	Public Property Get Height() : Height = height_ : End Property
	Public Property Let Height(vIn) : height_ = vIn : End Property		
	Public Property Get ColorDepth() : ColorDepth = Color_ : End Property
	Public Property Let ColorDepth(vIn) : Color_ = vIn : End Property		
	Public Property Get ImageType() : ImageType = ImageType_ : End Property
	Public Property Let ImageType(vIn) : ImageType_ = vIn : End Property		

	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':::  Functions to convert two bytes to a numeric value (long)   :::
	':::  (both little-endian and big-endian)                        :::
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	function lngConvert(strTemp)
		lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
	end function
	function lngConvert2(strTemp)
		lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
	end function

	function getDimenstion(strFile)
	   dim strPNG,strGIF,strBMP,strType
	   strType = ""
	   strImageType = "(unknown)"

	   strPNG = chr(137) & chr(80) & chr(78)
	   strGIF = "GIF"
	   strBMP = chr(66) & chr(77)
	   strType = Left(strFile,3)
	   if strType = strGIF then				' is GIF
	      strImageType = "GIF"
	      Width = lngConvert(Mid(strFile, 7, 2))
	      Height = lngConvert(Mid(strFile, 9, 2))
	   elseif left(strType, 2) = strBMP then		' is BMP
	      strImageType = "BMP"
	      Width = lngConvert(Mid(strFile, 19, 2))
	      Height = lngConvert(Mid(strFile, 23, 2))
	   elseif strType = strPNG then			' Is PNG
	      strImageType = "PNG"
	      Width = lngConvert2(Mid(strFile, 19, 2))
	      Height = lngConvert2(Mid(strFile, 23, 2))
	   else
	      'strBuff = flnm ''Mid(flnm, 0, -1)		' Get all bytes from file
	      lngSize = len(strFile)
	      flgFound = 0
	      strTarget = chr(255) & chr(216) & chr(255)
	      flgFound = instr(strFile, strTarget)
	      if flgFound = 0 then
	         exit function
	      end if
	      strImageType = "JPG"
	      lngPos = flgFound + 2
	      ExitLoop = False
	      do while (ExitLoop = False and lngPos < lngSize)
	         do while asc(mid(strFile, lngPos, 1)) = 255 and lngPos < lngSize
	            lngPos = lngPos + 1
	         loop
	         if asc(mid(strFile, lngPos, 1)) < 192 or asc(mid(strFile, lngPos, 1)) > 195 then
	            lngMarkerSize = lngConvert2(mid(strFile, lngPos + 1, 2))
	            lngPos = lngPos + lngMarkerSize  + 1
	         else
	            ExitLoop = True
	         end if
	     loop
	     if ExitLoop = False then
	        Width = -1
	        Height = -1
	     else
	        Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
	        Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
	     end if
	   end if
	end function
End Class
%>