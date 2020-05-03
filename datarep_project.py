"""
Project for Week 4 of "Python Data Representations".
Find differences in file contents.

Be sure to read the project description page for further information
about the expected behavior of the program.
"""
IDENTICAL = -1

def singleline_diff(line1, line2):
    """
    Inputs:
      line1 - first single line string
      line2 - second single line string
    Output:
      Returns the index where the first difference between
      line1 and line2 occurs.

      Returns IDENTICAL if the two lines are the same.
    """
    l1len = len(line1)
    l2len = len(line2)
    matchpos = IDENTICAL

    if line1 == line2:
        return IDENTICAL
    
    if l1len == 0 or l2len == 0:
        return 0
    
    if(l1len>=l2len):
        for item in range(l2len):
            if(line1[item] != line2[item]):
                matchpos = item
                break
            
            if(matchpos == IDENTICAL):
                matchpos = l2len
            
    elif(l2len>l1len):
        for item in range(l1len):
            if(line1[item] != line2[item]):
                matchpos = item
                break
            
            if(matchpos == IDENTICAL):
                matchpos = l1len      
                
    return matchpos

#print(singleline_diff(ln1, ln2))

def singleline_diff_format(line1, line2, idx):
    """
    Inputs:
      line1 - first single line string
      line2 - second single line string
      idx   - index at which to indicate difference
    Output:
      Returns a three line formatted string showing the location
      of the first difference between line1 and line2.

      If either input line contains a newline or carriage return,
      then returns an empty string.

      If idx is not a valid index, then returns an empty string.
    """    
    separator = ''
    formatted_string = ''
    len1 = len(line1)
    len2 = len(line2)
        
    #you do need to check whether the inputs are a single line this time.
    for item in range(len1):
        if(line1[item] == '\n' or line1[item] == '\r'):
            return formatted_string

    for item in range(len2):
        if(line2[item] == '\n' or line2[item] == '\r'):
            return formatted_string
        
    #return empty string if the lines are identical
    if(idx == IDENTICAL):
        return formatted_string
    
    if(len1 > len2 and idx != singleline_diff(line1,line2)):
        return formatted_string
    if(len2 > len1 and idx != singleline_diff(line1,line2)):
        return formatted_string
    
    separator = '{:{}>{}}'.format('^','=',idx+1)   
    
    formatted_string = line1 + '\n' + separator +  '\n' + line2 + '\n'
    
    return formatted_string
#print(repr(singleline_diff_format(ln1, ln2, singleline_diff(ln1, ln2))))

def multiline_diff(lines1, lines2):
    """
    Inputs:
      lines1 - list of single line strings
      lines2 - list of single line strings
    Output:
      Returns a tuple containing the line number (starting from 0) and
      the index in that line where the first difference between lines1
      and lines2 occurs.

      Returns (IDENTICAL, IDENTICAL) if the two lists are the same.
    """
    l1len = len(lines1)
    l2len = len(lines2)
    matchpos = idx = IDENTICAL
    
    if lines1 == lines2:
        return (IDENTICAL, IDENTICAL)
        
    if(l1len>=l2len):
        if (l2len > 0):
            for item in range(l2len):
                idx = singleline_diff(lines1[item],lines2[item])
                if(idx != IDENTICAL):
                    matchpos = item
                    break
                else:
                    if(matchpos != IDENTICAL):
                        idx = 0
                    else:
                        matchpos = l2len
        else:
            matchpos = idx = 0
                        
    elif(l2len>l1len):
        if (l1len > 0):
            for item in range(l1len):
                idx = singleline_diff(lines1[item],lines2[item])
                if(idx != IDENTICAL):
                    matchpos = item
                    break
                else:
                    if(matchpos != IDENTICAL):
                        idx = 0
                    else:
                        matchpos = l1len
        else:
            matchpos = idx = 0
       
    return(matchpos,idx)

#print(multiline_diff(list1,list2))

def get_file_lines(filename):
    """
    Inputs:
      filename - name of file to read
    Output:
      Returns a list of lines from the file named filename.  Each
      line will be a single line string with no newline ('\n') or
      return ('\r') characters.

      If the file does not exist or is not readable, then the
      behavior of this function is undefined.
    """
    stringlist = []
    
    #open file
    file_to_read = open(filename,'r',encoding='utf-8', errors='ignore')
    
      
    #create a list of line strings and remove \n from every line    
    for line in file_to_read.read().splitlines():     
        stringlist.append(line)        
  
    #close file
    file_to_read.close()

    #return the list of line strings
    return stringlist

#print(get_file_lines("project_test_file.txt"))

def file_diff_format(filename1, filename2):
    """
    Inputs:
      filename1 - name of first file
      filename2 - name of second file
    Output:
      Returns a four line string showing the location of the first
      difference between the two files named by the inputs.

      If the files are identical, the function instead returns the
      string "No differences\n".

      If either file does not exist or is not readable, then the
      behavior of this function is undefined.
    """

    line_number = index = IDENTICAL
    stringdiff = ''
    retstring = ''
    
    file1 = get_file_lines(filename1)
    file2 = get_file_lines(filename2)
    
    len1 = len(file1)
    len2 = len(file2)
        
    if(len1 < 0):
        file1 = ['']
        
    if(len2 < 0):
        file2 = ['']

    line_number, index = multiline_diff(file1,file2)
    
    if line_number == IDENTICAL and index == IDENTICAL:
        return "No differences\n"
    else:             
        if (index != 0):
            stringdiff = singleline_diff_format(file1[line_number],file2[line_number],index)
            retstring = 'Line '+str(line_number)+':\n'+ stringdiff
            if(len1 == 0 or len2 == 0):
                retstring+='\n'
        else: #index is identical          
            if(len1 > len2):
                retstring = 'Line '+str(line_number)+':\n'
                if (len2 == 0):
                    retstring += singleline_diff_format(file1[line_number],'',index)
                else:
                    retstring += singleline_diff_format(file1[line_number],file2[line_number],index)
            elif(len2 > len1):
                retstring = 'Line '+str(line_number)+':\n' 
                if (len1 == 0):
                    retstring += singleline_diff_format('',file2[line_number],index)+'\n'
                else:
                    retstring += singleline_diff_format(file1[line_number],file2[line_number],index)
                
    return retstring

#print(repr(file_diff_format("file8.txt", "file9.txt")))