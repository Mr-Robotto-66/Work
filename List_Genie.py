
def replace_character(list,replacechar):
    #replaces a given list of characters with a space
    return list.replace(replacechar,' ')

def find_dupes(list):
    #finds the duplicate values in the list and sorts them alphabetically
    return sorted([x for n, x in enumerate(list) if x in list[:n]])

def remove_dupes(list):
    #removes duplicates from the list and sorts the list alphabetically
    return sorted([x for n, x in enumerate(list) if x not in list[:n]])

def change_case(input,input_case):
	if input_case == 'All Upper Case':
		input = input.upper()
	elif input_case == 'All Lower Case':
		input = input.lower()
	return input

def format_list(list, list_format):
    formatted = list
    if list_format == 'Comma Delimited':
        formatted = ",".join(list)
    elif list_format == 'String Definition Query':
        formatted = "','".join(list)
        formatted = "('" + formatted + "')"
    elif list_format == 'Number Definition Query':
        formatted = ",".join(list)
        formatted = "(" + formatted + ")"
    elif list_format == 'Query Builder Format':
        formatted = " | ".join(list)
    return formatted

def find_matching(list1,list2):
	s = set(list2)
	matching = [x for x in list1 if x in s]
	return matching

def find_unmatched(list1,list2):
	s = set(list2)
	matching = [x for x in list1 if x not in s]
	return matching

def get_feature_list(layer,field):
    outer_list = []
    for row in arcpy.da.SearchCursor(layer,field):
##        arcpy.AddMessage("test")
        for item in row:
            item_replace = item.replace(" ","_")
            outer_list.append(unicode(item_replace))

##            outer_list.append(unicode(item))

    list = " ".join(outer_list)
    return list

if __name__ == '__main__':
    arcpy.AddMessage('===========================================')
    arcpy.AddMessage('     List Genie                   /)')
    arcpy.AddMessage('       made by           /\___/\ ((')
    arcpy.AddMessage("     Daniel Otto         \\'@_@'/  ))")
    arcpy.AddMessage('daniel.otto@gov.bc.ca    {_:Y:.}_//')
    arcpy.AddMessage("------------------------{_}^-'{_}----------")
    arcpy.AddMessage('===========================================')

    import arcpy,sys

    list1 = sys.argv[1]# The first user input to turn into a list
    list2 = sys.argv[3] # The second user input to turn into a list, optional
    replace_char = sys.argv[7] # The list of characters to remove from the lists (eg.",;^$@*^)
    list_format = sys.argv[6] # the format choices to turn the list into(Comma Delimited, String Definition Query, Number Definition Query)
    list_case = sys.argv[5] # does the user want to turn all the upper or lower case (All Upper Case, All Lower Case)

    field1 = sys.argv[2]
    field2 = sys.argv[4]

    if field1 is not '#':
        list1 = get_feature_list(list1,field1)
    if field2 is not '#':
        list2 = get_feature_list(list2,field2)


    # Change the case of the list if the user wants
    list1 = change_case(list1,list_case)
    # Replace the characters specficied in the argument with spaces
    for char in replace_char:
    	list1 = replace_character(list1,char)
    #split the string into a list at the spaces
    list1split = list1.split()
    #Call the find dupes function to find all the duplicate vales
    dupes1 = find_dupes(list1split)
    #If there is more than 0 duplicates return the duplicates
    if len(dupes1) > 0:
    	arcpy.AddMessage(str(len(dupes1)) + ' duplicate value(s) in list 1:')
    	arcpy.AddMessage(' ')
    	for dupe in dupes1:
    		arcpy.AddMessage(dupe)
    else:
    	arcpy.AddMessage('There are no duplicate values in list 1')
    arcpy.AddMessage('-------------------------------------------')
    arcpy.AddMessage(list2)
    # return the list to the user with duplicates removed
    no_dupes1 = remove_dupes(list1split)
    arcpy.AddMessage (str(len(no_dupes1))+ ' unique items in list 1:')
    arcpy.AddMessage(' ')
    for no_dupe in no_dupes1:
    	arcpy.AddMessage(no_dupe)
    arcpy.AddMessage('-------------------------------------------')
    #format the list as the user specifies
    arcpy.AddMessage('Here is your list1 formatted in the style you picked:')
    arcpy.AddMessage(' ')
    arcpy.AddMessage(format_list(no_dupes1,list_format))
    arcpy.AddMessage('--------------------------------------------')


    if list2 is not '#':
    	list2 = change_case(list2,list_case)
    	for char in replace_char:
    		list2 = replace_character(list2,char)
    	list2split = list2.split()
    	dupes2 = find_dupes(list2split)
    	if len(dupes2) > 0:
    		arcpy.AddMessage(str(len(dupes2)) + ' duplicate value(s) in list 2:')
    		arcpy.AddMessage(' ')
    		for dupe in dupes2:
    			arcpy.AddMessage(dupe)
    	else:
    		arcpy.AddMessage('There are no duplicate values in list 2')
    	arcpy.AddMessage('-----------------------------------------')
    	no_dupes2 = remove_dupes(list2split)
    	arcpy.AddMessage (str(len(no_dupes2))+ ' unique items in list2:')
    	arcpy.AddMessage(' ')
    	for no_dupe in no_dupes2:
    		arcpy.AddMessage(no_dupe)
    	arcpy.AddMessage('------------------------------------------')
    	#format the list as the user specifies
    	arcpy.AddMessage('Here is your list2 formatted in the style you picked:')
    	arcpy.AddMessage(' ')
    	arcpy.AddMessage(format_list(no_dupes2,list_format))

    	arcpy.AddMessage(' ')
    	arcpy.AddMessage('==========================================')
    	arcpy.AddMessage('    |\_/|      Comparing      D\___/\\')
    	arcpy.AddMessage('    (0_0)         The          (0_o)')
    	arcpy.AddMessage('   ==(Y)==      Lists...        (V)')
    	arcpy.AddMessage('--(u)---(u)-----------------oOo--U--oOo--')
    	arcpy.AddMessage('__|_______|_______|_______|_______|______')
    	arcpy.AddMessage(' ')

    	# find matching between lists
    	matching = find_matching(no_dupes1,no_dupes2)
    	arcpy.AddMessage(str(len(matching))+ ' matching values between the list 1 and list 2:')
    	arcpy.AddMessage(' ')
    	for matched in matching:
    		arcpy.AddMessage(matched)
    	arcpy.AddMessage('-------------------------------------------')
    	# find unmatched between the 2 lists
    	unmatched1 = find_unmatched(no_dupes1,no_dupes2)
    	arcpy.AddMessage(str(len(unmatched1))+ ' values in list 1 not in list 2:')
    	arcpy.AddMessage(' ')
    	for unmatched in unmatched1:
    		arcpy.AddMessage(unmatched)
    	arcpy.AddMessage('-------------------------------------------')
    	unmatched2 = find_unmatched(no_dupes2,no_dupes1)
    	arcpy.AddMessage(str(len(unmatched2))+ ' values in list 2 not in list 1:')
    	arcpy.AddMessage(' ')
    	for unmatched in unmatched2:
    		arcpy.AddMessage(unmatched)
    	arcpy.AddMessage('===========================================')









