'Kamil L	anda 12/2022
'https://github.com/KamilLanda/TreeSort
'licence: 	'CC0 1.0 Universal (CC0 1.0) Public Domain Dedication https://creativecommons.org/publicdomain/zero/1.0/
				'CC0 1.0 Univerzální (CC0 1.0) Potvrzení o statusu volného díla https://creativecommons.org/publicdomain/zero/1.0/deed.cs

option explicit

private gpi() 'array with the indexed letters
private gi&, gpOut() 'index and array for result

Sub TreeSortSimple
	on local error goto bug
	dim oDoc as object
	dim p(), i&, j&, k&, s$, s1$, min&, max&, p1(), iCount&, a&, index&, iLen&, iEmpties&
	max=0 : min=&hffff& 'maximal and minimal Ascii code for used characters (min=&h10ffff& for Full-Unicode)
	dim pb(max to min) as boolean 'info array, which characters are used
	
	rem small Czech alphabet - edit it for your use, separate the characters by space, the space has reserved word: space
	'! but don't add characters with more letters like 'ch', this version is only for one-letter characters!
	'!!! there must be all characters used in data for this Simple version!!!
	CONST sINIT="space a á b c č d ď e é ě f g h i í j k l m n ň o ó p q r ř s š t ť u ů ú v w x y ý z ž"
	p1=split(sINIT, " ")
	for i=lbound(p1) to ubound(p1)
		if p1(i)="space" then
			p1(i)=" "
			exit for
		end if
	next i
	
	rem the array with items to sort :-) -> use only characters that are in constant sINIT!
	p=array("aa", "baba", "ba", "oba", "baobab", "ob", "bob", "boa", "a", "aa", "aaa", "ah", "ahoj", "ahu", "aha", "ah", "aa", "", "", "grázl", "aa", "aab", "cáco", "cácora", " chamraď", "čmuchači", "chechot", "vochechule", "vochmelka", "smraďoch", "svoloč", "tupec", "jasoň", "buřič", "drsoň", "blivajz", "lemra", "čumič", "žvanil", "grázl", "jasoň", "buřič", "drsoň", "bli vajz", "blivajz", "lemra", "čumič", "bab", "cab", "bábovka", "bab", "cab", "babizna", "aaa", "aa", "bb", "aa", "a a", "aa", " aaa", "aa", "aa", "aaa", "aa", "a", "b", "bb")

rem 1) Detect used characters
	if ubound(p)=-1 then exit sub 'entry input array
	for i=lbound(p) to ubound(p)
		s=p(i) 'current word
		if s<>"" then
			iLen=len(s)
			for j=1 to iLen
				a=asc(mid(s, j, 1)) 'ascii code of current character
				if a<min then min=a 'minimal ascii code
				if a>max then max=a 'maximal ascii code
				pb(a)=true 'character is used
			next j
		else 'count of empty strings
			iEmpties=iEmpties+1
		end if
	next i

rem 2) create indexed arrays
	redim gpi(1+ubound(p1)) : i=0
	dim p2(min to max) 'the index is the ascii code, the value is the sorting order
	for each s in p1
		a=Asc(s)
		if a>=min AND a<=max then 'maybe the character is used
			if pb(a) then
				iCount=iCount+1
				gpi(iCount)=s 'character is used
				p2(a)=iCount
			end if
		end if
	next
	redim preserve gpi(iCount)
	'xray gpi
	'xray p2

rem 3) build the tree
	dim tree(ubound(gpi)) : tree(0)=0 'first element is: array(Count of word; true/false = if the current branch has some sub-branches)
	dim level as variant, p3() 'level is current branch; p3() is temporary array
	for i=lbound(p) to ubound(p)
		s=p(i) : iLen=len(s) 'current word
		level=tree 'current branch for loop
		for j=1 to iLen 'go through characters in word
			s1=mid(s, j, 1) 'current character
			rem add the array to the tree
			index=p2(asc(s1))
			if isEmpty(level(index)) then 'still no character in cell in branch
				if j=iLen then 'last character from word, so it means end of word
					level(index)=1 'put only number
					tree(0)=tree(0)+1 'increase the count of all words
				else 'no last character
					redim p3(iCount) 'new branch with empty values
					p3(0)=0 'count is 0; but the sub-branch is
					level(index)=p3 'add new branch to tree
					level=level(index) 'take new branch
				end if
			elseif isArray(level(index)) then 'the letter is in branch
				if j=iLen then 'last character from word, so it means end of word
					level(index)(0)=level(index)(0)+1 'increase the count of word
					tree(0)=tree(0)+1 'increase the count of all words
				else 'no last character
					level=level(index) 'take new branch
				end if
			else 'the letter was last for some previous word, so move the number of count to new sub-array to index(0)
				redim p3(iCount) 'new branch with empty values
				p3(0)=level(index) 'count is moved to new branch
				level(index)=p3 'add new branch to tree
				level=level(index) 'take new branch		
			end if	
		next j
	next i

rem 4) read the tree
	'xray(tree)
	redim gpOut(tree(0)) 'same size for output array like the input array
	a=msgbox("Yes - list without duplicities" & chr(13) & "No - only uniques", 35)
	if a=6 then 'Yes - LIST
		gi=0
		recurArray(tree, "") 'condition in recursion is >0
		gpOut(0)=array("TOTAL", tree(0)) 'total count of items		
	elseif a=7 then 'No - UNIQUES
		gi=1
		for i=1+lbound(tree) to ubound(tree) 'for uniques is problem with 1st index tree(0) because it is total count of all items, so it need go from tree(1)
			if NOT isEmpty(tree(i)) then recurArrayUniques(tree(i), gpi(i)) 'condition in this recursion is =1
		next i
		gpOut(0)=array("UNIQUES", gi-1) 'total count of items
	else 'Cancel
		exit sub
	end if
	redim preserve gpOut(gi-1)

rem 5)	report the result
	s=""
	for i=lbound(gpOut) to ubound(gpOut)
		s=s & gpOut(i)(0) & " " & gpOut(i)(1) & "×"
		if i<> ubound(gpOut) then s=s & chr(13)
	next i
	msgbox(s, 0, "Sorted & Counted")
	exit sub
bug:
	bug(Erl, Err, Error, "TreeSort")
End Sub

Function recurArray(level as variant, s$) 'recursion for reading the branches, output to global array gpOut
	on local error goto bug
	dim i&
	if level(0)>0 then 'for the count of current word
		gpOut(gi)=array( s, level(0) ) 'add to global output array
		gi=gi+1
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

Function recurArrayUniques(level as variant, s$) 'only one difference unto recurArray is in condition If ... Then
	on local error goto bug
	dim i&
	if level(0)=1 then 'for the count of current word
		gpOut(gi)=array( s, level(0) ) 'add to global output array
		gi=gi+1
	end if
	if isArray(level) then
		for i=1+lbound(level) to ubound(level) 'next indexes in branch
			if NOT isEmpty(level(i)) then recurArrayUniques(level(i), s & gpi(i)) 'get the sub-branch
		next i
	end if
	exit function
bug:
	bug(Erl, Err, Error, "recurArrayUniques")
End Function

Sub bug(sErl$, sErr$, sError$, sFce$) 'error message
	msgbox("line: " & sErl & chr(13) & sErr & ": " & sError, 16, sFce)
	stop
End Sub

