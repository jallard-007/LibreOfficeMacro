REM  *****  BASIC  *****

Rem opens the given file in hidden mode
function OpenDocAsHidden(URL as string) as object
	dim FileProperties(0) as New com.sun.star.beans.PropertyValue
	FileProperties(0).name = "Hidden"
	FileProperties(0).value = True
	OpenDocAsHidden = StarDesktop.loadComponentFromURL(Url, "_blank", 0, FileProperties())
	
	Rem Closes the file, just clean up
	Rem Doc.Close(False)
end function

Rem gets the current directory (directory that ThisComponent is saved in)
function GetCurrentDirectory
	sUrl = ThisComponent.getURL()
	if sUrl = "" then
		GetCurrentDirectory = ""
		exit function
	end if
	sParts = Split(sUrl, "/")
	redim Preserve sParts(0 to UBound(sParts) - 1)
	GetCurrentDirectory = Join(sParts, "/")
end function

Rem returns an array of file names matching certain file types in the given directory
function GetFilesInDirectory(directory as string) as Variant
	Rem file types to look for
	dim fileTypes(3) as string 
	fileTypes(0) = "*.ods"
	fileTypes(1) = "*.csv"
	fileTypes(2) = "*.xlsx"
	fileTypes(3) = "*.xls"
	
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
		sValue = Dir$(directory + "/" + fileType,0)
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

Rem return type for loadSections function below
type loadSectionsReturnType
	sections as variant
	sectionsSize as integer
	costCodeRowIndex as integer
end type

Rem loads information regarding column headers
function loadSections(Sheet as variant) as variant
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
		Rem display the error to the user
		msgbox "Error: Could not find column headings (No cell containing """ & headerID & """ found within the search range A1 to A" _
			& numRowsToSearch & ")" & CHR(13) & CHR(13) & _
			"To fix this error, insert the text """ & headerID & """ within the specified range above, on the same row as your column headings"
		Rem return an integer to signify that there was an error
		dim error1 as integer
		loadSections = error1
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
	dim currFolder as string
	currFolder = GetCurrentDirectory
	if currFolder = "" then
		msgbox "Error: Please save this file, and then try again"
		exit sub
	end if

	dim refDoc as object
	refDoc = OpenDocAsHidden(currFolder & "/file2.ods")
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
		msgbox "Nothing to import"
		refDoc.Close(True)
		exit sub
	end if

	dim data(sectionsInfo.sectionsSize) as variant
	for column = 0 to sectionsInfo.sectionsSize
		data(column) = loadSectionOfEstimate(sheet, sectionsInfo.costCodeRowIndex, column + 1)
	next

	refDoc.Close(True)
	addToDoc(sectionsInfo, data, "file2")
end sub

sub addToDoc(sectionInfo as loadSectionsReturnType, data as variant, fileName as string)
	Rem dim costCodeDoc as object
	Rem dim currFolder as string
	Rem currFolder = GetCurrentDirectory
	Rem if currFolder = "" then
	Rem 	exit sub
	Rem end if
	Rem costCodeDoc = OpenDocAsHidden(currFolder & "/CostCodes.ods")

	dim sheet as object
	sheet = ThisComponent.Sheets(0)
	dim result as variant
	result = loadSections(sheet)

	if VarType(result) = V_INTEGER then
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
	dim offset as integer
	offset = column
	for each section in data
		sheet.columns.insertByIndex(column,1)
		row = result.costCodeRowIndex
		sheet.getCellByPosition(column,row).setString("&" + fileName + "-" + sectionInfo.sections(column - offset))
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
