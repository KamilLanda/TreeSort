'Kamil L	anda 12/2022
'https://github.com/KamilLanda/TreeSort
'licence: 	'CC0 1.0 Universal (CC0 1.0) Public Domain Dedication https://creativecommons.org/publicdomain/zero/1.0/
				'CC0 1.0 Univerzální (CC0 1.0) Potvrzení o statusu volného díla https://creativecommons.org/publicdomain/zero/1.0/deed.cs

option explicit
	
private gpi() 'array with the indexed letters
private gj&, gi&, gpOut() 'index and array for result
private goStatusbar as object, gc2&, giStep& : const giStepLimit=40 'for progressbar; giStepLimit is count of steps for progressbar actualization

Sub TreeSort
	on local error goto bug
	dim oDoc as object
	dim p(), i&, j&, k&, sInit$, s$, s1$, s2$, min&, max&, p1(), p2(), ptemp(), iCount&, a&, a1&, a2&, index&, iLen&, iEmpties&, min2&, max2&
	dim b2letter as boolean, b1 as boolean, p2letters(), pUnknown(), iUnknown&, i2letter&, sUnknown$
	dim time0 as date, time as date, sTime$, sTime1$, sTime2$, sTime3, sTime4$
	time0=Now 'total time
	max=0 : min=&hffff& : min2=min : max2=max 'maximal and minimal Ascii code for used characters for UTF-8 (for Full-Unicode: min=&h10ffff&)
	dim pb(max to min) as integer, pb2() 'info array, which characters are used
	const iLimit=100 'limit to show the result in Calc (>); or show it only in msgbox (<=)
	
	const bLowerCase=FALSE 'Hewn case-insensitive sorting; TRUE=transform all to lowercases
	
	rem data from TXT file
	'p=loadFileString("d:\temp\cs.dic") 'url, [encoding (default: utf-8)], [delimiter (default: Enter - that is gotten from 1st line in file automatically)]
	rem data from Array()
	p=array("ana", "analýza", "analýze", "CHa", "chamraď", "cha", "ChAMRAĎ", "barbora", "buč", "Cd", "cab", "asc", "líná", "Ɯ", "líné", "ab", "ƉƂ", "žába", "čicHací", "CHeCHot", "voChechule", "vochmelka", "smraďoch", "svoloč", "tupec", "grázl", "jasoň", "drsoň", "blivajz", "lemra", "čumič", "GRÁZL", "JASOŇ", "BUŘIČ", "DRSOŇ", "BLIVAJZ", "LEMRA", "ČUMIČ", "BAB", "CAB", "BÁBOVKA", "bab", "cab", "bab", "aaa", "aa", "bb", "aa", "aaa", "aa", "aaa", "aa", "aa", "aaa", "aa", "a", "b", "bb", "bdd")
	
	resetStatusBar
	time=Now 'measure the time
	goStatusbar=getStatusController.StatusIndicator
	goStatusbar.start("Items: " & ubound(p), 3*ubound(p)) 'initialization of progress bar
	
	rem Czech alphabet with space and numbers - edit it for your use, separate the characters by space; the original space has reserved word: space; double Quotation is "" (standard mark for Quotation in basic strings)
	sINIT="space 0 1 2 3 4 5 6 7 8 9 a A á Á b B c C č Č d D ď Ď e E é É ě Ě f F g G h H ch cH Ch CH i I í Í j J k K l L m M n N ň Ň o O ó Ó p P q Q r R ř Ř s S š Š t T ť Ť u U ů Ů ú Ú v V w W x X y Y ý Ý z Z ž Ž"
	'sINIT="a b c ch cH " & chr(101) & " d" 'you can also use this for some character, but do proper separation with spaces in string and use reserved word 'space' for space :-)
	if bLowerCase then sInit=LCase(sInit)
	p1=split(trim(sINIT), " ") 'get the array from constant
	for i=lbound(p1) to ubound(p1)
		if p1(i)="" then
			msgbox "Empty space at position: " & i & chr(13) & "At the back of character: " & p1(i-1)
			exit sub
		elseif p1(i)="space" then
			p1(i)=" "
		end if
	next i

rem 0) prepare the array with supposed characters from sINIT
	goStatusbar.setText("Prepare the supposed characters")
	max=32 : min=382
	gj=-1 : min2=min : max2=max
	redim p2(max to min)
	for i=lbound(p1) to ubound(p1)
		s=p1(i)
		a1=Asc(s)
		if a1<min then min=a1
		if a1>max then max=a1
		j=i+1
		if Len(s)=1 then 'only one character
			if isEmpty(p2(a1)) then 'no index
				p2(a1)=gj 'index of the character as normal array
				gj=gj-1
			elseif p2(a1)(0)<>0 then 'there is some index so current character is duplicate
				if NOT bLowerCase then msg("Duplicate in constant sINIT: " & chr(13) & chr(a1) & chr(13) & "At position: " & j)  'control for duplicities
			else 'there is empty array
				p2(a1)(0)=gj 'set index for single char to array(0)
				gj=gj-1
			end if
		else 'two-letters character
			s2=Mid(s, 2, 1) : a2=Asc(s2) '2nd letter
			if isEmpty(p2(a1)) then '1st letter isn't in array, so it isn't two-letters character
				redim ptemp(a2 to a2)
				ptemp(a2)=gj
				gj=gj-1
				p2(a1)=array(0, ptemp) 'set array as index
				if a1<lbound(p2Letters) OR a1>ubound(p2letters) then 'repair bounds for two-letters array
preskok:
					if a1<min2 then min2=a1
					if a1>max2 then max2=a1
					redim preserve p2letters(min2 to max2)
					p2letters(a1)=ptemp
				end if
			elseif isArray(p2(a1)) then 'there is already some two-letter
				ptemp=p2(a1)(1)
				b1=false 'for repairing bound for two-letters array
				if a2=lbound(ptemp) OR a2=ubound(ptemp) then 'control for duplicities
					if NOT bLowerCase then msg("Duplicate in constant sINIT: " & chr(13) & chr(a1) & s2 & chr(13) & "At position: " & j)
				end if
				if a2<lbound(ptemp) then 'repair left bound
					redim preserve ptemp(a2 to ubound(ptemp))
					b1=true
				end if
				if a2>ubound(ptemp) then 'repair right bound
					redim preserve ptemp(lbound(ptemp) to a2)
					b1=true
				end if
				ptemp(a2)=gj
				gj=gj-1
				p2(a1)=array(p2(a1)(0), ptemp) 'set array as index
				if b1 then 'repair bounds for two-letters array
					GOTO preskok 'I know, 21th century is :-) -> but for example it is easy; thereunto you can try to create more optimalized example without GoTo :-)
				end if
			else 'there is only single char
				redim ptemp(a2 to a2)
				ptemp(a2)=gj
				gj=gj-1
				p2(a1)=array(p2(a1), ptemp) 'set array as index
			end if
		end if
	next i
	redim preserve p2(min to max) 'truncate array to used bounds

rem 1) detect really used characters
	goStatusbar.setText("Detect really used characters")
	min2=min : max2=max
	redim pUnknown(min to max)
	for i=lbound(p) to ubound(p)
		updateProgressbar(i)
		s=p(i) 'current word
		if s="" then 'empty line in TXT file
			iEmpties=iEmpties+1
		else 'some word
			if bLowerCase then
				s=LCase(s)
				p(i)=s
			end if
			for j=1 to len(s)
				s1=mid(s, j, 1) '1st letter
				a1=Asc(s1) 'ascii of 1st letter
semUnk:
				if a1<min2 OR a1>max2 then 'lower or higher unknown letter
					if a1<min then
						min=a1
						redim preserve p2(min to max)
						redim preserve pUnknown(min to max)
					elseif a1>max then
						max=a1
						redim preserve p2(min to max)
						redim preserve pUnknown(min to max)
					end if
					pUnknown(a1)=true
					p2(a1)=ABS(p2(a1)) 'mark the character as used
				else 'known letter
					if NOT isArray(p2(a1)) then 'only one character
						if NOT isEmpty(p2(a1)) then
							p2(a1)=ABS(p2(a1)) 'mark the character as used
						else
							pUnknown(a1)=true 'add unknown character
						end if
					else 'two-letters character
						if j<len(s) then '1st letter isn't last letter
							s2=Mid(s, j+1, 1) : a2=Asc(s2) '2nd letter
							if a2<lbound(p2(a1)(1)) OR a2>ubound(p2(a1)(1)) OR isEmpty(p2(a2)) then '2nd letter is unknown
								if p2(a1)(0)=0 then
									p2(a1)(0)=ABS(gj) '1st letter isn't part of two-letter and it is unknown letter, mark one as used
									pUnknown(a1)=true 'add unknown character
									gj=gj-1
								elseif p2(a2)(0)=0 then '2nd letter is unknown
									p2(a2)(0)=ABS(gj) 'mark the character as used
									pUnknown(a1)=true 'add unknown character
									gj=gj-1
								else
									p2(a1)(0)=ABS(p2(a1)(0)) '1st letter is normal, mark one as used
									a1=a2
									j=j+1
									b1=true
									GOTO semUnk
								end if
							else '2nd letter is normal
								j=j+1
								if isEmpty(p2(a1)(1)(a2)) then '2nd letter is normal letter
									p2(a1)(0)=ABS(p2(a1)(0)) '1st letter is normal
									if p2(a2)(0)=0 then '2nd letter is unknown but in known part
										p2(a2)(0)=ABS(gj) 'mark the character as used
										pUnknown(a1)=true 'add unknown character
										gj=gj-1
									else '2nd letter is normal
										p2(a2)=ABS(p2(a2)) 'mark the character as used
									end if
								else
									p2(a1)(1)(a2)=ABS(p2(a1)(1)(a2)) '2nd letter is 2nd part of two-letter
								end if
							end if
						else '1st letter is last letter
							p2(a1)(0)=ABS(p2(a1)(0)) 'mark the character as used
						end if
					end if
				end if
			next j
		end if
	next i
	redim preserve p2(min to max) 'truncate array to used bounds
	j=0
	for i=lbound(pUnknown) to ubound(pUnknown) 'get the count of unknown chracters
		if pUnknown(i)=true then j=j+1
	next i
	redim gpi(ubound(p1)+j+ubound(pUnknown)) 'gpi is array for recursion
	recurGpi(p2) 'array with the tree with indexes to sorted characters
	redim p1(gi+j)
	j=0
	for i=lbound(gpi) to ubound(gpi) 'get used characters from known characters
		if NOT isEmpty(gpi(i)) then
			j=j+1
			p1(j)=gpi(i)
		end if
	next i
	for i=lbound(pUnknown) to ubound(pUnknown) 'get used characters from unknown characters
		if NOT isEmpty(pUnknown(i)) then
			j=j+1
			s=chr(i)
			p1(j)=s
			sUnknown=sUnknown & " " & s
		end if
	next i
	rem message about unknown characters
'	if NOT isEmpty(pUnknown) then 'warning for unknown characters
'		if 1<>msgbox("Unknown characters:" & chr(13) & sUnknown, 49, "Continue?") then exit sub
'	end if
	rem array p1(): there are only really used sorted characters instead the supposed characters at start from sINIT, index(0) is reserved for count of words

rem 2) create indexes for really used characters
	time=Now-time : sTime1=Minute(time) & ":" & Second(time) : time=Now 'time of Indexing
	goStatusbar.setText("Create indexes")
	min2=min : max2=max : gj=1 'gj is count of used characters 
	redim p2(min to max)
	redim ptemp()
	for i=1+lbound(p1) to ubound(p1)
		s=p1(i)
		a=Asc(s)
		j=i+1
		if Len(s)=1 then 'only one character
			if isEmpty(p2(a)) then 'no index
				p2(a)=gj 'index of character as normal array
				gj=gj+1
			else 'there is empty array
				p2(a)(0)=gj 'set index for single char to array(0)
				gj=gj+1
			end if
		else 'two-letters character
			s2=Mid(s, 2, 1) : a2=Asc(s2) '2nd letter
			if isEmpty(p2(a)) then '1st letter isn't in array, so it isn't two-letters character
				redim ptemp(a2 to a2)
				ptemp(a2)=gj
				gj=gj+1
				p2(a)=array(0, ptemp) 'set array as index
				if a<lbound(p2Letters) OR a>ubound(p2letters) then 'repair bounds for two-letters array
					if a<min2 then min2=a
					if a>max2 then max2=a
					redim preserve p2letters(min2 to max2)
					p2letters(a)=ptemp
				end if
			elseif isArray(p2(a)) then 'there is already some two-letter
				ptemp=p2(a)(1)
				b1=false 'for repairing bound for two-letters array
				if a2<lbound(ptemp) then 'repair left bound
					redim preserve ptemp(a2 to ubound(ptemp))
					b1=true
				end if
				if a2>ubound(ptemp) then 'repair right bound
					redim preserve ptemp(lbound(ptemp) to a2)
					b1=true
				end if
				ptemp(a2)=gj
				gj=gj+1
				p2(a)=array(p2(a)(0), ptemp) 'set array as index
				if b1 then 'repair bounds for two-letter array
					if a<min2 then min2=a
					if a>max2 then max2=a
					redim preserve p2letters(min2 to max2) 'repair bounds
					p2letters(a)=ptemp
				end if
			else 'there is only single char
				redim ptemp(a2 to a2)
				ptemp(a2)=gj
				gj=gj+1
				p2(a)=array(p2(a), ptemp) 'set array as index
			end if
		end if
	next i
	redim preserve p2(min to max) 'truncate array to used bounds
	rem array p2(): is structure with "pointers" from Ascii codes to array p1() with sorted really used characters
	rem so the "transformation filter" including two-letters characters is ready :-)

rem 3) build the tree
	time=Now-time : sTime2=Minute(time) & ":" & Second(time) : time=Now 'time of Indexing
	goStatusbar.setText("Building the tree")
	b2letter=false : iCount=ubound(p1)
	dim tree(ubound(p1)) : a=0 : tree(0)=a 'index (0) is Longint for total count of items in tree
	dim level as variant, p3() 'level is current branch; p3() is temporary array
	for i=lbound(p) to ubound(p)
		updateProgressbar(ubound(p)+i)
		s=p(i) : iLen=len(s) 'current word
		level=tree 'current branch for loop
		for j=1 to iLen 'go throught characters in word
			s1=mid(s, j, 1) 'current character
			a1=Asc(s1)
			rem get the index of character
			if NOT isArray(p2(a1)) then 'normal character
				index=p2(a1)
				b2letter=false
			elseif j<len(s) then 'test for two-letter
				s2=mid(s, j+1, 1)
				a2=Asc(s2)
				if a2>=lbound(p2(a1)(1)) AND a2<=ubound(p2(a1)(1)) then 'test for two-letter
					if NOT isEmpty(p2(a1)(1)(a2)) then 'it is two-letter
						s1=s1 & s2
						index=p2(a1)(1)(a2)
						b2letter=true
					else 'normal character
						b2letter=false
						GOTO pridat
					end if
				else '2nd character is normal letter
pridat:
					index=p2(a1)(0)
					b2letter=false
					level=addToTree(tree, level, index, j, iLen, b2letter, iCount)
					j=j+1
					a1=a2
					index=p2(a1)(0)
				end if
			else 'last letter
				index=p2(a1)(0)
			end if
			rem put character to tree
			level=addToTree(tree, level, index, j, iLen, b2letter, iCount)
			if b2letter then j=j+1 'increase for two-letters character
		next j
	next i
	rem tree is builded
	gpi=p1 'use global array for for recursion

rem 4) read the tree
	time=Now-time : sTime3=Minute(time) & ":" & Second(time) : time=Now 'time of Building
	goStatusbar.setText("Reading is faster :-)")
	gc2=2*ubound(p) 'for progressbar in recurArray()
	gi=0 : redim gpOut(tree(0)) 'same size for output array like the input array
	recurArray(tree, "") 'set result to array() gpOut
	redim preserve gpOut(gi-1) : gpOut(0)(0)="TOTAL" 'small optimalization of array
	s="" 'for final msgbox

rem 5) report the result
	if tree(0)+iEmpties<>1+ubound(p) then msgbox("input array: " & 1+ubound(p) & chr(13) & "tree: " & tree(0), 16, "Bad Count of items") 'small control for count of items
	time=Now-time : sTime4=Minute(time) & ":" & Second(time) 'time of Reading
	dim oRange as object
	if ubound(p)>iLimit then 'many items, show ones in Calc
		oDoc=StarDesktop.LoadComponentFromUrl("private:factory/scalc", "_blank", 0, array())
		oDoc.lockControllers
		oDoc.addActionLock
		oRange=oDoc.Sheets(0).getCellRangeByPosition(0, 0, 1, ubound(gpOut))
		oRange.setDataArray(gpOut)
		oDoc.unlockControllers
	else 'only few items, show msgbox
		for i=lbound(gpOut) to ubound(gpOut)
			s=s & gpOut(i)(0) & " " & gpOut(i)(1) & "×" & chr(13)
		next i
	end if
	resetStatusbar
sem:
	time0=Now-time0 'total time
	if ubound(p)>iLimit then sTime="Detection: " & sTime1 & chr(13) & "Indexing: " & 	sTime2 & chr(13) & "Building: " & sTime3 & chr(13) & "Reading: " & sTime4 & _
		chr(13) & "TOTAL: " & Minute(time0) & ":" & Second(time0) & chr(13) & chr(13) 'report the time only for many items
	msgbox(sTime & s)', 0, tree(0) & " items")
	exit sub
bug:
	bug(Erl, Err, Error, "TreeSort")
End Sub

Function recurArray(level as variant, s$) 'recursion for reading the branches, output to ARRAY
	on local error goto bug
	dim i&
	if level(0)>0 then 'index(0) is for the count of current word
		gpOut(gi)=array(s, CLng(level(0)) ) 'add to global output array
		gi=gi+1
		updateProgressbar(gc2+gi)
	end if
	if isArray(level) then
		for i=1+lbound(level) to ubound(level) 'next indexes in branch
			if NOT isEmpty(level(i)) then recurArray(level(i), s & gpi(i)) 'get the sub-branch
		next i
	end if
	exit function
bug:
	bug(Erl, Err, Error, "recurArray")
End Function

Function addToTree(tree(), level(), index&, j&, iLen&, b2letter as boolean, iCount&) as array 'add item to the tree
	on local error goto bug
	dim p3()
	if isEmpty(level(index)) then 'still no character in cell in branch
		if isLastChar(j, iLen, b2letter) then 'last character from word, so it means end of word
			tree(0)=tree(0)+1 'increase the count of all words
			level(index)=1 'it is 1st occurence of word
		else 'no last character
			redim p3(iCount)
			level(index)=p3 'add new branch to tree
			level=level(index) 'take new branch
		end if
	else 'the letter is in branch
		if isLastChar(j, iLen, b2letter) then 'last character from word, so it means end of word
			level(index)(0)=level(index)(0)+1 'increase the count of word
			tree(0)=tree(0)+1 'increase the count of all words
		elseif NOT isArray(level(index)) then 'no last character, but there was only end of previous word so change the element to array
			redim p3(iCount)
			p3(0)=level(index) 'count is moved to new branch
			level(index)=p3 'add new branch to tree
			level=level(index) 'take new branch
		else
			level=level(index) 'take new branch
		end if
	end if
	addToTree=level
	exit function
bug:
	bug(Erl, Err, Error, "addToTree")
End Function

Function recurGpi(p(), optional s$) 'recursion to count the order of really used characters
	on local error goto bug
	if isMissing(s) then s=""
	dim i&
	for i=lbound(p) to ubound(p)
		if i=97 then
			wait 10
		end if
		if NOT isEmpty(p(i)) then
			if NOT isArray(p(i)) then 'normal letter
				if p(i)>0 then 'letter is used
					gpi(p(i))=s & chr(i)
					gi=gi+1
				end if
			else 'two-letters character
				if p(i)(0)>0 then '1st letter is used as normal letters
					gpi(p(i)(0))=s & chr(i)
					gi=gi+1
					recurGpi(p(i)(1), chr(i))
				elseif p(i)(0)<0 then '1st letter is used only as part of two-letter
					recurGpi(p(i)(1), chr(i))
				end if
			end if
		end if
	next i
	exit function
bug:
	bug(Erl, Err, Error, "recurGpi ")
End Function

Function isLastChar(j&, iLen&, b2letter as boolean) 'test for last character in word
	on local error goto bug
	isLastChar=iif( j=iLen OR (j=iLen-1 AND b2letter), true, false) 'tested character is at final position or it is last 'ch'
	exit function
bug:
	bug(Erl, Err, Error, "isLastChar")
End Function

Function loadFileString(sUrl$, optional sEncoding$, optional sEnter$) as array 'read file to array
	on local error goto bug
	if isMissing(sEncoding) then sEncoding="UTF-8" 'default encoding is UTF-8
	dim s$, oSfa as object, oTextStream as object, oStream as object, s1$
	oSfa=CreateUNOService("com.sun.star.ucb.SimpleFileAccess")
	if oSfa.exists(sUrl) then 'the file exists
		oStream=oSfa.openFileRead(sUrl)
		oTextStream=CreateUNOService("com.sun.star.io.TextInputStream")
		oTextStream.InputStream=oStream
		if (Trim(sEncoding)<>"") then oTextStream.Encoding=sEncoding
		rem detect the Enter in file
		if isMissing(sEnter) then 'read 1st line to detect Enter later
			s1=loadFileLine(sUrl, sEncoding) 'read 1st line from the file
			if s1="" then goto chyba 'file doesn't exists or maybe it is empty file
		end if
		s=oTextStream.readString(array(), false) 'read all text from file with split() later is faster than read the file line by line
		oTextStream.closeInput
		oStream.closeInput
	else 'the file doesn't exist
chyba:
		msgbox(sUrl & " doesn't exists", 16, "Problem with input file")
		stop
	end if
	if isMissing(sEnter) then 'detect Enter from 1st read line
		dim i&, s2$
		i=Len(s1)
		s2=mid(s, i+1, 1) 'next character behind the 1st read line
		if s2=chr(10) then 'linux
			sEnter=chr(10)
		elseif s2=chr(13) then 'win or mac
			s2=mid(s, i+2, 1)
			if s2=chr(10) then sEnter=chr(13) & chr(10) else sEnter=chr(13)
		end if
	end if
	loadFileString=split(s, sEnter) 'split the text from file by Enter
	exit function
bug:
	bug(Err, Erl, Error, "loadFileString")
End Function

Function loadFileLine(sUrl$, sEncoding$) as string 'read one line to help with detection of Enter
	on local error goto bug
	dim s$, oSfa as object, oTextStream as object, oStream as object, s1$
	oSfa=CreateUNOService("com.sun.star.ucb.SimpleFileAccess")
	if oSfa.exists(sUrl) then 'the file exists
		oStream=oSfa.openFileRead(sUrl)
		oTextStream=CreateUNOService("com.sun.star.io.TextInputStream")
		oTextStream.InputStream=oStream
		if (Trim(sEncoding)<>"") then oTextStream.Encoding=sEncoding
		s=oTextStream.readLine 'read 1st line from file
		oTextStream.closeInput
		oStream.closeInput
		loadFileLine=s
	else 'the file doesn't exist
		loadFileLine=""
	end if
	exit function
bug:
	bug(Erl, Err, Error, "loadFileLine")
End Function

Function getStatusController() as object 'get oDoc.CurrentController for progressbar, also if only Basic Editor IDE is running
	on local error goto bug
	dim oDoc as object
	oDoc=ThisComponent
	if isNull(CreateUnoService("com.sun.star.reflection.CoreReflection").getType(oDoc)) OR _
	isNull(CreateUnoService("com.sun.star.script.Invocation").createInstanceWithArguments(array(oDoc))) then 'regular oDoc isn't, try to get other Libre window like Basic editor
		dim oComponents as object, oComponent as object
		oComponents=StarDesktop.Components.CreateEnumeration
		while oComponents.hasMoreElements 'try all Libre windows
			oComponent=oComponents.NextElement
			if oComponent.Identifier="com.sun.star.script.BasicIDE" then 'Basic Editor is active
				getStatusController=oComponent.CurrentController
				exit function
			end if
		wend
		oDoc	=StarDesktop.LoadComponentFromUrl("private:factory/swriter", "_blank", 0, array()) 'open new document in Writer for progressbar
	else
		getStatusController=oDoc.CurrentController
	end if
	exit function
bug:
	bug(Erl, Err, Error, "getStatusController")
End Function

Sub resetStatusbar 'only reset a statusbar
	on local error goto bug
	if isNull(goStatusbar) then goStatusbar=getStatusController.StatusIndicator
	goStatusbar.end
	goStatusbar.reset	
	exit sub
bug:
	bug(Erl, Err, Error, "resetStatusbar")
End Sub

Sub updateProgressbar(i&) 'update the progressbar after some steps (faster than one by one)
	on local error goto bug
	giStep=giStep+1
	if giStep=giStepLimit then 'set new value
		goStatusbar.setValue(i)
		giStep=0
	end if
	exit sub
bug:
	bug(Erl, Err, Error, "updateProgressbar")
End Sub

Sub bug(sErl$, gErr$, gError$, sFce$) 'error message
	msgbox("line: " & sErl & chr(13) & gErr & ": " & gError, 16, sFce)
	stop
End Sub

Sub msg(s$) 'show only msgbox nad stop the macro
	msgbox(s, 48, "Error during process")
	stop
End Sub

