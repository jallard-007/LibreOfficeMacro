REM  *****  BASIC  *****

function OpenDocAsHidden(URL as string) as object
	dim FileProperties(0) as New com.sun.star.beans.PropertyValue
	FileProperties(0).name = "Hidden"
	FileProperties(0).value = True
	OpenDocAsHidden = StarDesktop.loadComponentFromURL(Url, "_blank", 0, FileProperties())
	
	Rem Closes the file, just clean up
	Rem Doc.Close(False)
end function

function GetCurrentFolder
	sUrl = ThisComponent.getURL()
	sParts = Split(sUrl, "/")
	redim Preserve sParts(0 to UBound(sParts) - 1)
	GetCurrentFolder = Join(sParts, "/")
end function

function GetFilesInCurrDirectory() as Variant
	Rem file types to look for
	dim fileTypes(3) as string 
	fileTypes(0) = "*.ods"
	fileTypes(1) = "*.csv"
	fileTypes(2) = "*.xlsx"
	fileTypes(3) = "*.xls"

	Rem get current folder to search in
	dim sPath as string
	sPath = GetCurrentFolder
	
	Rem Get current file, we don't include it
	sUrl = ThisComponent.getURL()
	urlSplit = Split(sUrl, "/")
	currFile = urlSplit(UBound(urlSplit))

	Rem initialize array to store files, starting size of capacity (actually capacity + 1 since basic is weird)
	dim capacity as integer
	capacity = 5
	dim files(capacity) as string
	dim numOfFiles as integer
	numOfFiles = 0

	dim sValue as string
	for each fileType In fileTypes
		sValue = Dir$(sPath + getPathSeparator + fileType,0)
		do until sValue = ""
			if sValue <> currFile then
				if numOfFiles > capacity then
					capacity = capacity * 2
					Redim preserve files(0 to capacity)
				end if
				files(numOfFiles) = sValue
				numOfFiles = numOfFiles + 1
			end if
			sValue = Dir$
		loop
	next
	
	if numOfFiles > 0 then 
		redim preserve files(0 to numOfFiles - 1)
	end if

	GetFilesInCurrDirectory = files
	Rem displays all found files
	Rem for each t in GetFilesInCurrDirectory
	Rem msgbox t
	Rem next
end function

type loadSectionsReturnType
	sections as variant
	sectionsSize as integer
	costCodeRowIndex as integer
end type

function loadSections(Sheet as variant) as loadSectionsReturnType
	Rem find row of columns headings (section names)
	Rem we look for the first row that has something in its leftmost column, up to numRowsToSearch row
	const headerID as string = "Cost Code"
	dim row as integer
	const numRowsToSearch as integer = 10
	dim cellValue as string
	for row = 0 to (numRowsToSearch - 1)
		cellValue = Sheet.getCellByPosition(0,row).getString()
		if cellValue = headerID then
			exit for
		end if
	next
	
	Rem check that we found data within the limit
	if row = numRowsToSearch then 
		Rem return an integer to signify that there was an error
		dim error1 as integer
		loadSectionOfEstimate = error1
		Rem display the error to the user
		msgbox "Error: Could not find column headings (No cell containing """ & headerID & """ found within the search range A1 to A" _
			& numRowsToSearch & ")" & CHR(13) & CHR(13) & _
			"To fix this error, insert the text """ & headerID & """ within the specified range, on the same row as your column headings"
		exit function
	end if
	
	dim sectionsSize as integer
	sectionsSize = 5
	dim sections(sectionsSize) as string
	dim sectionsCount as integer
	sectionsCount = 0
	dim column as integer
	column = 1
	dim cellContent as string
	cellContent = Sheet.getCellByPosition(sectionsCount + 1,row).getString()
	do while cellContent <> "" and (LEFT(cellContent, 1) <> "&")
		if sectionsSize < sectionsCount then
			sectionsSize = sectionsSize * 2
			redim preserve sections(0 to sectionsSize)
		end if
		sections(sectionsCount) = cellContent
		sectionsCount = sectionsCount + 1
		cellContent = Sheet.getCellByPosition(sectionsCount + 1,row).getString()
	loop
	
	if sectionsCount > 0 then 
		redim preserve sections(0 to sectionsCount - 1)
	end if
	dim returnObject as loadSectionsReturnType
	returnObject.sections = sections
	returnObject.sectionsSize = sectionsCount - 1
	returnObject.costCodeRowIndex = row
	loadSections = returnObject
end function

sub main
	dim refDoc as object
	refDoc = OpenDocAsHidden(GetCurrentFolder & "/file2.ods")
	dim sheet as object
	sheet = refDoc.Sheets(0)
	dim sectionsInfo as variant
	sectionsInfo = loadSections(sheet)
	
	if VarType(sectionsInfo) = V_INTEGER then
		Rem failed
		refDoc.Close(True)
		exit sub
	end if

	if sectionsInfo.sectionsSize < 0 then
		Rem no sections found
		refDoc.Close(True)
		exit sub
	end if

	dim data(sectionsInfo.sectionsSize) as variant
	for column = 0 to sectionsInfo.sectionsSize
		data(column) = loadSectionOfEstimate(sheet, sectionsInfo.costCodeRowIndex, column + 1)
	next

	refDoc.Close(True)
	addToDoc(sectionsInfo, data)
end sub

sub addToDoc(sectionInfo as loadSectionsReturnType, data as variant)
	Rem dim costCodeDoc as object
	Rem costCodeDoc = OpenDocAsHidden(GetCurrentFolder & "/CostCodes.ods")

	dim sheet as object
	sheet = ThisComponent.Sheets(0)
	dim result as loadSectionsReturnType
	result = loadSections(sheet)

	if VarType(sectionsInfo) = V_INTEGER then
		Rem failed
		exit sub
	end if

	Rem find first empty column along headerID row
	dim row as integer
	row = result.costCodeRowIndex
	dim column as integer
	column = result.sectionsSize + 2
	dim cellContent as string
	cellContent = sheet.getCellByPosition(column,row).getString()
	do while cellContent <> ""
		column = column + 1
		cellContent = sheet.getCellByPosition(column,row).getString()
	loop
	for each section in data
		sheet.columns.insertByIndex(column,1)
		row = result.costCodeRowIndex
		sheet.getCellByPosition(column,row).setString("&")
		row = row + 1
		cellContent = sheet.getCellByPosition(0,row).getString()
		do while cellContent <> ""
			if section.containsKey(cellContent) then
				sheet.getCellByPosition(column,row).setString(section.get(cellContent))
			end if

			row = row + 1
			cellContent = sheet.getCellByPosition(0,row).getString()
		loop
		column = column + 1
	next
end sub

function loadSectionOfEstimate(Sheet as variant, startingRow as integer, column as integer) as variant
	dim map as variant
	map = com.sun.star.container.EnumerableMap.create("string", "string")
	dim row as integer
	row = startingRow + 1
	dim costCode as string
	costCode = Sheet.getCellByPosition(0,row).getString()
	do while costCode <> ""
		cellContent = Sheet.getCellByPosition(column,row).getString()
		map.put(costCode, cellContent)
		row = row + 1
		costCode = Sheet.getCellByPosition(0,row).getString()
	loop
	loadSectionOfEstimate = map
end function
