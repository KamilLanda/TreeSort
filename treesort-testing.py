# Kamil Landa 12/2022
# https://github.com/KamilLanda/TreeSort
# licence: 	'CC0 1.0 Universal (CC0 1.0) Public Domain Dedication https://creativecommons.org/publicdomain/zero/1.0/
# 			'CC0 1.0 Univerzální (CC0 1.0) Potvrzení o statusu volného díla https://creativecommons.org/publicdomain/zero/1.0/deed.cs

# run function main()

import time

# array with sorted letters
gpi=[0, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
gi=0
gpOut=[None]
icount=len(gpi)

# array with indexes for Ascii codes
p2=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]

# timsort from https://realpython.com/sorting-algorithms-python
def merge(left, right):
    if len(left) == 0:
        return right

    if len(right) == 0:
        return left

    result = []
    index_left = index_right = 0

    while len(result) < len(left) + len(right):
        if left[index_left] <= right[index_right]:
            result.append(left[index_left])
            index_left += 1
        else:
            result.append(right[index_right])
            index_right += 1

        if index_right == len(right):
            result += left[index_left:]
            break

        if index_left == len(left):
            result += right[index_right:]
            break

    return result

def insertion_sort(array, left=0, right=None):
    if right is None:
        right = len(array) - 1

    for i in range(left + 1, right + 1):
        key_item = array[i]
        j = i - 1
        while j >= left and array[j] > key_item:
            array[j + 1] = array[j]
            j -= 1

        array[j + 1] = key_item

    return array

def timsort(array):
    min_run = 32
    n = len(array)
    for i in range(0, n, min_run):
        insertion_sort(array, i, min((i + min_run - 1), n - 1))

    size = min_run
    while size < n:
        for start in range(0, n, size * 2):
            midpoint = start + size - 1
            end = min((start + size * 2 - 1), (n-1))
            merged_array = merge(left=array[start:midpoint + 1], right=array[midpoint + 1:end + 1])
            array[start:start + len(merged_array)] = merged_array

        size *= 2

    return array
# end timsort


def treesort(p):
    global gpi, gi, gpOut, p2, icount

# 3) build the tree
    startTime=time.time()
    print("building the tree")
    tree=[None]*len(gpi)
    tree[0]=0
    for x in p:
        s=x # current word
        ilen=len(s)
        ilen2=ilen-1
        level=tree # current branch for loop
        j=0
        while j<ilen: # go through characters in word
            s1=s[j]
            # add array to the tree
            index=p2[ord(s1)]

            if ( isinstance( level[index], type(None) ) ): # still no character in cell in branch
                if (j==ilen2): # last character from word, so it means end of word
                    level[index]=1 # put only number
                    tree[0]+=1 # increase the count of all words
                else:
                    p3=[None]*icount # new branch with empty values
                    p3[0]=0 # count is 0; but the sub-branch is
                    level[index]=p3 # add new branch to tree
                    level=level[index] # take new branch
            elif ( isinstance( level[index], list) ): # the letter is in branch
                if (j==ilen2): # last character from word, so it means end of word
                    level[index][0]+=1 # increase the count of word
                    tree[0]+=1 # increase the count of all words
                else:
                    level=level[index]
            else:
                p3=[None]*icount # new branch with empty values
                p3[0]=level[index] # count is moved to new branch
                level[index]=p3 # add new branch to tree
                level=level[index] # take new branch

            j+=1 #next j

# 4) read the tree
    print(str(time.time() - startTime))
    startTime=time.time()
    print("reading the tree")
    a=tree[0]+1
    gpOut=[None]*a
    recurArray(tree, "")
    gpOut[0]=["TOTAL", tree[0]]
    gpOut=gpOut[0:gi]

# 5) write the result
    print(str(time.time() - startTime))
    #print(gpOut)


def recurArray(level, s):
    global gpOut, gi
    i=0
    if ( isinstance( level, int) ):
        gpOut[gi]=[s, level]
        gi+=1
    elif (level[0]>0):
        gpOut[gi]=[s, level[0]]
        gi+=1
    if ( isinstance( level, list) ):
        i=1
        while i<len(level):
            if ( not isinstance( level[i], type(None) ) ):
                recurArray(level[i], s + gpi[i])
            i+=1 # next i


# change the variable c, but maximum value is 26 if you don't change the gpi[] and p2[]; I tested with 20 in editor Pyzo -> it took about 6 minutes and 13GB RAM!!!
def generate_big_array(c): #create big array from 6letters words containing the characters from gpi[]
    global gpi
    d=c**6
    print("generating the big array: " + str(d) + " items")
    p=[0]*d
    i=c
    o=0
    while i>0:
        j=c
        while j>0:
            k=c
            while k>0:
                l=c
                while l>0:
                    m=c
                    while m>0:
                        n=c
                        while n>0:
                            p[o]=gpi[i] + gpi[j] + gpi[k] + gpi[l] + gpi[m] + gpi[n]
                            o+=1
                            n-=1
                        m-=1
                    l-=1
                k-=1
            j-=1
        i-=1
    return p

def standard_sort(p): # standard python sorting
    sorted(p)

#data from TXT file
def loadFileString(sUrl):
  with open(sUrl, encoding="utf-8",) as f:
    ss=f.read()
    f.close()
    return ss.splitlines()

def main():
    # set the variable p[] from standard array; or from file; or from generate_big_array()
    #p=["AA", "BABA", "BA", "OBA", "BAOBAB", "OB", "BOB", "BOA", "A", "AA", "AAA", "AH", "AHOJ", "AHU", "AHA", "AH", "AA", "GRAZL", "AA", "AAB", "CACO", "CACORA", "CHAMRAD", "CMUCHACI", "CHECHOT", "VOCHECHULE", "VOCHMELKA", "SMRADOCH", "SVOLOC", "TUPEC", "JASON", "BURIC", "DRSON", "BLIVAJZ", "LEMRA", "CUMIC", "ZVANIL", "GRAZL", "JASON", "BURIC", "DRSON", "BLIVAJZ", "BLIVAJZ", "LEMRA", "CUMIC", "BAB", "CAB", "BABOVKA", "BAB", "CAB", "BABIZNA", "AAA", "AA", "BB", "AA", "AA", "AA", "AAA", "AA", "AA", "AAA", "AA", "A", "B", "BB"]

    #p=loadFileString("d:/temp/cs.txt") # !!! BE SURE THERE ARE ONLY CHARACTERS DEFINED IN gpi[] AND INDEXED IN p2[] !!!

    p=generate_big_array(8) #parameter: count of characters from gpi[]: 1=only A; 2=A,B; 3=A,B,C; 4=A,B,C,D etc. ! maximum is 26 that means about 300 millions items (26^6); I tested with 20 in editor Pyzo -> it took about 6 minutes and 13GB RAM!!!
    print("big array created")

    startTime=time.time()
    size=len(p)
    print("timsort is running")
    p=timsort(p)
    print("timsort sorted: " + str(time.time() - startTime))
    startTime=time.time()

    # print("internal python sorting")
    # standard_sort(p)
    # print("internal sorted: " + str(time.time() - startTime))
    # startTime=time.time()

    print("treesort is running")
    treesort(p)
    print("treesort sorted: " + str(time.time() - startTime))

main()