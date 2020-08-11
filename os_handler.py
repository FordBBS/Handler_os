####################################################################################################
#                                                                                                  #
# Python, OS handler, BBS					   						                               #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/08/08, BBS:	- Move to BBS modules
# 2020/08/09, BBS: 	- Implemeneted 'IUser_create_file_fromstr'
#
#***************************************************************************************************

#*** Function Group List ***************************************************************************
# - Constants & Important parameters
# - Timestamp
# - File & Directory handler
# - Content Comparison
# - Text file parser



#*** Library Import ********************************************************************************
# Operating system
import  datetime
import  shutil
import  pathlib
import  os
import  binascii
import  json

# BBS_PyGem
import  dataset_handler 	as hs_dataset
from  	dataset_handler 	import hs_prep_AnyList
from  	dataset_handler 	import hs_prep_StrList 



#*** Function Group: Constants & Important parameters **********************************************
def getconst_chr_quote():
	RetVal = chr(34)
	return RetVal

def getconst_chr_linebreak():
	RetVal = [chr(10), chr(13)]
	return RetVal

def getconst_chr_path():
	RetVal = [chr(47), chr(92)] 		# ["/", "\"]
	return RetVal



#*** Function Group: Timestamp *********************************************************************
def IBase_get_timestamp(format_str):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a timestamp in requested format

	[str] format_str, target time format, 'yyyymmdd_hhmm' is used if required format is invalid
	
	'''

	#*** Initialization ****************************************************************************
	list_format	= ["yyyymmdd_hhmmss","yyyymmdd_hhmm","yymmdd_hhmmss","yymmdd_hhmm", "yyyymmdd", \
	 				"yymmdd", "hhmmss", "y/m/d hh:mm:ss"]
	list_syntax	= ["%Y%m%d_%H%M%S","%Y%m%d_%H%M%S","%y%m%d_%H%M%S","%y%m%d_%H%M", "%Y%m%d", \
					"%y%m%d", "%H%M%S", "%Y/%b/%d %H:%M:%S"]

	#*** Operations ********************************************************************************
	#--- Post Input Validation ---------------------------------------------------------------------
	if not(isinstance(format_str,str)): format_str = list_format[2]
	format_str	= format_str.lower()

	#--- Gets the date-time format based on requested format ---------------------------------------
	if not (format_str in list_format): format_str = list_syntax[2]
	else: format_str = list_syntax[list_format.index(format_str)]

	#--- Gets current timestamp --------------------------------------------------------------------
	RetVal	= datetime.datetime.now()
	RetVal	= RetVal.strftime(format_str)

	#--- Release -----------------------------------------------------------------------------------
	return RetVal

def IBase_get_formatted_date(strTimestamp, format_month, format_year, flg_dash):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a current timestamp-date into the requested format
		or
		Convert the requested timestamp, strTimestamp, to a timestamp-date in the requested format

	[str]  strTimestamp,	timestamp value to be converted (yyyymmdd),
							uses current timestamp if this is empty
	[int]  format_month,	Flag of target format for month, 1: Full Name (May), 
							2: Number (5), Others: 3-Char-Abbreviation (Jun, Jan)
	[int]  format_year,		Flag of target format for year,  1: 2-Digit, Others: 4-Digit
	[bool] flg_dash,		Flag of Using "-" as a conjunction
	
	'''

	#*** Initialization ****************************************************************************
	strTimestamp = str(strTimestamp)
	strTimestamp = strTimestamp.replace(" ", "")
	strTimestamp = strTimestamp.replace("-", "")

	if len(strTimestamp) == 0:   strTimestamp = IBase_get_timestamp("yyyymmdd")
	elif len(strTimestamp) == 6: strTimestamp = IBase_get_timestamp("yyyymmdd")[:2] + strTimestamp

	# Format Selection : Day
	tar_fm_day = "%d" 								# Default, 01, 18, 25

	# Format Selection : Month
	tar_fm_month = "%b"								# Default, Abbreviation
	if format_month == 1: 	tar_fm_month = "%B"		# Full name : May, November, August
	elif format_month == 2: tar_fm_month = "%m"		# Number

	# Format Selection : Year
	tar_fm_year = "%Y"								# Default, 4-digit year
	if format_year == 1: 	tar_fm_year = "%y" 		# Short version, (20)19, (20)20

	# Combines Formats
	if not isinstance(flg_dash, bool): flg_dash = False
	if flg_dash: str_format = chr_space = "-"
	else: chr_space = " "

	str_format = tar_fm_day + chr_space + tar_fm_month + chr_space + tar_fm_year

	#*** Operations ********************************************************************************
	ts_year  = int(strTimestamp[:4])
	ts_month = int(strTimestamp[4:][:2])
	ts_date	 = int(strTimestamp[6:])
	sys_date = datetime.datetime(ts_year, ts_month, ts_date)

	#--- Release -----------------------------------------------------------------------------------
	return sys_date.strftime(str_format)

def IBase_get_formatted_time(strTimestamp, format_hour, flg_colon):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a current timestamp-time into the requested format
		or 
		Convert the requested timestamp, strTimestamp, to a timestamp-time in the requested format

	[str]  strTimestamp, timestamp value to be converted (hhmmss),
						 uses current timestamp if this is empty
	[int]  format_hour,	 Flag of target format for hour,  1: 00-12 with AM/PM, Others: 00 - 23
	[bool] flg_colon,	 Flag of Using ": as a conjunction
	
	'''

	#*** Initialization ****************************************************************************
	if len(strTimestamp) == 0 or len(strTimestamp) != 6: strTimestamp = IBase_get_timestamp("hhmmss")

	# Format Selection : Hour
	tar_fm_hr = "%H" 								# Default, 00-23
	if format_hour == 1: tar_fm_hr = "%I%p" 		# 00-12 with AM/PM symbol

	# Format Selection : Minute
	tar_fm_min = "%M"								# Default, 00-59

	# Format Selection : Second
	tar_fm_sec = "%S"								# Default, 00-59

	# Combines Formats
	if not isinstance(flg_colon, bool): flg_colon = False
	if flg_colon: str_format = chr_space = ":"
	else: chr_space = " "

	str_format 	= tar_fm_hr + chr_space + tar_fm_min + chr_space + tar_fm_sec

	#*** Operations ********************************************************************************
	cur_time 	= IBase_get_timestamp("yyyymmdd")
	ts_year 	= int(cur_time[:4])
	ts_month	= int(cur_time[4:][:2])
	ts_date		= int(cur_time[6:])

	ts_hour		= int(strTimestamp[:2])
	ts_min		= int(strTimestamp[2:][:2])
	ts_sec		= int(strTimestamp[4:])
	sys_date 	= datetime.datetime(ts_year, ts_month, ts_date, ts_hour, ts_min, ts_sec)

	#--- Release -----------------------------------------------------------------------------------
	return sys_date.strftime(str_format)



#*** Function Group: Content Comparison ************************************************************
def IBase_read_binary(pathFile):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of binary content of 'pathFile'. "NA" is returned if file doesn't exist

	[str] pathFile, A string of target file's path

	'''

	#*** Input Validation **************************************************************************
	if not os.path.exists(str(pathFile)): return "NA"

	#*** Operations ********************************************************************************
	with open(pathFile, "rb") as tmpbin: return list(binascii.hexlify(tmpbin.read()))	

def IBase_compare_list_issame(listA, listB):
	#*** Documentation *****************************************************************************
	'''Documentation

		return a boolean flag value whether 2 input lists are same
		True: 	'listA' and 'listB' are same
		False: 	'listA' and 'listB' are different

	[list] listA, Binary content list A
	[list] listB, Binary content list B
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listA, list) and isinstance(listB, list)): return 101

	#*** Operations ********************************************************************************
	flg_same = False

	if len(listA) == len(listB):
		for idx_chr in range(0, len(listA)):
			if listA[idx_chr] != listB[idx_chr]:
				break
		else: flg_same = True
	return flg_same



#*** Function Group: File & Directory handler ******************************************************
def IBase_get_desktop_path():
	#*** Documentation *****************************************************************************
	'''Documentation

		A path of Dekstop is returned if it's possible. Otherwise, CWD's path is returned.

	'''
	try: 	return os.path.expanduser("~/Desktop")
	except: return ""

def IBase_get_name_on_path_by_level(strPath, intLevel):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a name of folder at specific level 'intLevel' in 'strPath'
		If 'intLevel' is larger then actual level of 'strPath', last file or folder's name is returned

		e.g. strPath  = "Formulas/res_gmd/rep_bag_wait_time_$(gmd)#.frm/type2"
			 Result: "rep_bag_wait_time_$(gmd)#.frm" (intLevel = 3, intLevel = -2)
			 Result: "type2: (intLevel >= 4)

	[str] strPath,  string path to be referred
	[int] intLevel, Target level to get the name
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(strPath, str) and len(strPath) > 0): return strPath
	if not isinstance(intLevel, int) or intLevel == 0: 		return strPath
	
	#*** Initialization ****************************************************************************
	chr_path = getconst_chr_path()[1]
	strPath  = IBase_get_formatted_path(strPath, chr_path)
	amnt_lvl = strPath.count(chr_path) + 1

	#*** Operations ********************************************************************************
	#--- Correct minus value on 'intLevel' ---------------------------------------------------------
	if intLevel < 0: intLevel = amnt_lvl + intLevel + 1
	if intLevel > amnt_lvl or intLevel < 0: intLevel = amnt_lvl

	#--- Get name of folder/file at target level ---------------------------------------------------
	return strPath.split(chr_path)[intLevel - 1]

def IBase_get_splitted_strpath(strPath):
	#*** Documentation *****************************************************************************
	'''Documentation

		return a list consists of 3 information; FlgExist, FolderPath, FileName, FileSuffix
		FlgExist:   True if strPath does exist, false if strPath does not exist
		FolderPath: Path of deepest folder in strPath
		FileName:   Name of file include file extension, empty if strPath is a directory path
		FileSuffix: File Extension, empty if strPath is a directory path

	[str] strPath, string path to be verified and splitted
	
	'''

	#*** Input Validation **************************************************************************
	strPath = str(strPath)
	if len(strPath) == 0: return ["", ""]

	#*** Initialization ****************************************************************************
	listRes = []

	#*** Operations ********************************************************************************
	listRes.append(os.path.exists(strPath))
	tmpSuffix = os.path.splitext(strPath)[1]

	if os.path.isfile(strPath) or (not listRes[0] and len(tmpSuffix) > 0):
		listRes.extend(os.path.split(strPath))
		listRes.append(tmpSuffix)
	
	elif os.path.isdir(strPath) or len(tmpSuffix) == 0:
		listRes.extend([strPath, "", ""])

	return listRes

def IBase_get_splitted_path_byword(listPath, keyword, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of each path in 'listPath' splitted by 'keyword'
		Any path that doesn't contain 'keyword' will be returned as [ThatPath, ""]
		
		e.g. "ema/dataprocessing/Formulas/EU/UN-ECE_R120_Rev1" with keyword = "Formulas" results in
			["ema/dataprocessing/Formulas", "EU/UN-ECE_R120_Rev1"]

			 "ema/dataprocessing/Formulas/EU" with keyword = "Region" results in
			["ema/dataprocessing/Formulas/EU", ""]

	[list] listPath, A list of path to be splitted by 'keyword'
	[str]  keyword,  A string used as a marker for spliting. 'keyword' folder will be the last
					 folder for the first path of splitted path

	[bool] flg_case, True: string case is considered differently, False: string case is ignored
	'''

	#*** Input Validation **************************************************************************
	listPath = hs_prep_StrList(listPath)

	try: keyword = str(keyword)
	except: keyword = ""

	if len(listPath) == 0 or not (isinstance(keyword, str) and len(keyword) > 0): return [[], []]

	#*** Initialization ****************************************************************************
	if not isinstance(flg_case, bool): flg_case = True
	if not flg_case: keyword = keyword.lower()

	chr_path = getconst_chr_path()[1]
	listRes  = []

	for idx in range(0, len(listPath)): listRes.append(["", ""])

	#*** Operations ********************************************************************************
	listPath = [IBase_get_formatted_path(x, chr_path) for x in listPath]

	for idx_main, eachPath in enumerate(listPath):
		if flg_case: effPath = eachPath
		else: effPath = eachPath.lower()

		if not keyword in effPath:
			listRes[idx_main][0] = eachPath
		else:
			idx = effPath.find(keyword) + len(keyword)
			listRes[idx_main][0] = eachPath[:idx]
			listRes[idx_main][1] = eachPath[idx + 1:] 

	return listRes

def IBase_get_formatted_path(strPath, chr_conj):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a formatted string path

	[str] strPath,  target string-path to be formatted
	[str] chr_conj, target character to be used as a conjuction of each folder
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(strPath, str)): 	 return ""
	if len(strPath) < 2:				 return ""
	if not (isinstance(chr_conj, str)):  chr_conj = ""

	#*** Initialization ****************************************************************************
	list_chr_ToBeConverted = getconst_chr_path()

	if len(chr_conj) == 0: chr_conj = list_chr_ToBeConverted[0]

	#*** Operations ********************************************************************************
	for eachChar in list_chr_ToBeConverted: strPath = strPath.replace(eachChar, chr_conj)
	if strPath[len(strPath) - 1:] == chr_conj: strPath = strPath[:len(strPath) - 1]
	return strPath

def IBase_getlist_content(strPath, flg_onlyfile):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of content under strPath directory, if strPath is a file, it will be rollback
		to its folder and all contents under that folder will be read

	[str] strPath,  target directory to be walking through
	[bool] flg_clr_redundant, True: Remove redundant directory, False: Disbaled Removing process
	[bool] flg_onlyfile, True: Only file will be listed, False: Empty folder will be listed too

	'''

	#*** Input Validation **************************************************************************
	strPath = str(strPath)
	if not os.path.exists(strPath): return []

	#*** Initialization **************************************************************************** 
	if not isinstance(flg_onlyfile, bool): flg_onlyfile = False

	chr_path = getconst_chr_path()[1]
	strPath  = IBase_get_formatted_path(strPath, chr_path)
	listRes  = []

	#*** Operations ********************************************************************************
	#--- Get complete directory list ---------------------------------------------------------------
	for listRoot, listFolder, listFile in os.walk(strPath):
		if len(listFolder) == 0 and len(listFile) == 0 and not flg_onlyfile: listRes.append(listRoot)
		for eachFile in listFile:
			curPath = os.path.join(listRoot, eachFile)
			listRes.append(curPath)

	listRes = hs_dataset.IBase_get_sorted_list([listRes], "", "")[0]
	
	#--- Release -----------------------------------------------------------------------------------
	return listRes

def IBase_getlist_content_filter(listPath, listFilter, flg_onlyfile, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a list of content under all path in 'listPath' directory. If any path is a file, it
		will be fallback to its folder and all contents under that folder will be read.
		Filtering can be used by providing multi-column filter 'listFilter'

		Result Format
		[listFullPath]

	[list] listPath,     A list of target path
	[list] listFilter, 	 A multi-column list contain filter elements. Each internal list will be used
						 for 'Some-Met' filtering process (One tree met among combos returns Passed)
						 e.g. [["EU", "US"], "all", ["leg_gmd", "leg_ecs", "leg_fuel"], ...]
	[bool] flg_onlyfile, True: Only file will be listed, False: Empty folder will be listed too
	[bool] flg_case, 	 True: string case is considered differently, False: string case is ignored

	'''

	#*** Input Validation **************************************************************************
	listPath = hs_prep_StrList(listPath)

	if len(listPath) == 0 or (len(listPath) == 1 and len(listPath[0]) == 0): return []

	#*** Initialization ****************************************************************************
	if not isinstance(flg_onlyfile, bool): flg_onlyfile = True
	if not isinstance(flg_case, bool):     flg_case     = True
	listRes = []

	#*** Operations ********************************************************************************
	#--- Content Listing ---------------------------------------------------------------------------
	for eachPath in listPath:
		tmpContent = IBase_getlist_content(eachPath, flg_onlyfile)
		listVal    = [0]*len(tmpContent)

		for idx, eachContent in enumerate(tmpContent):
			if hs_dataset.IBase_filter_str_multilist(eachContent, listFilter, 1, True, flg_case):
				listVal[idx] = 1

		listRes.extend(hs_dataset.IBase_get_reduced_list(tmpContent, listVal))

	#--- Release -----------------------------------------------------------------------------------
	if len(listRes) > 0: listRes = hs_dataset.IBase_get_sorted_list([listRes], [0], [1])[0]
	return listRes

def IBase_getlist_diff_content(listContA, listContB, flg_comp, listRootPath):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of content that only 'listContA' has. Listing-only-folder can be selected
		as option via flg_mode. 

	[list] listContA,    Base Content List
	[list] listContB,    Comparitor Content List
	[bool] flg_comp,     True: Enabled content comparison, False: Disabled content comparison
	[list] listRootPath, a list consists of rootpath of 'listContA' and rootpath of 'listContB'

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listContA, list) and isinstance(listContB, list)): return listContA

	#*** Initialization **************************************************************************** 
	if not isinstance(flg_comp, bool): flg_comp = False
	if not (isinstance(listRootPath, list) and len(listRootPath) == 2): listRootPath = ["", ""]

	rootPathA  = listRootPath[0]
	rootPathB  = listRootPath[1]
	listRes    = []
	listCommon = []

	#*** Operations ********************************************************************************
	#--- Get rootPath if they are not provided -----------------------------------------------------
	if len(rootPathA) == 0: print("TODO: Get a root path among list of paths")
	if len(rootPathB) == 0: print("TODO")

	#--- Disabled Comparison: Subtraction Basis ----------------------------------------------------
	setB 	= set([x.replace(rootPathB, rootPathA) for x in listContB])
	listRes = [x for x in listContA if x not in setB]

	if not flg_comp: return listRes	

	#--- Enabled Comaprison: Binary comparison Basis -----------------------------------------------
	listCommon = list(set(listContA) - set(listRes))

	for eachItem in listCommon:
		if os.path.isfile(eachItem):
			curTarRefPath = eachItem.replace(rootPathA, rootPathB)

			if not os.path.exists(curTarRefPath): listRes.append(eachItem)
			else:
				binBase  = open(eachItem, "rb")						# Read in Binary mode
				readBase = list(binascii.hexlify(binBase.read()))	# Bytes Object
				binBase.close()

				binRef   = open(curTarRefPath, "rb")
				readRef  = list(binascii.hexlify(binRef.read()))
				binRef.close()

				flg_same = IBase_compare_list_issame(readBase, readRef)
				if not flg_same: listRes.append(eachItem)

	return hs_dataset.IBase_get_sorted_list([listRes], [0], [1])[0]

def IBase_getlist_method_content(listFilePath, flg_issorted):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of consists of listMethod and listFilePath where method means the content
		unique number of that file when there are more than one file exist in 'listFilePath'

	[list] listFilePath, A list of file path to be processed
	[bool] flg_issorted, True:  'listFilePath' has been sorted, 
						 False: 'listFilePath' hasn't been sorted. Sort will be applied

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listFilePath, list)): return 101
	if len(listFilePath) == 0: 				 return listFilePath

	#*** Initialization ****************************************************************************
	listMethod = []
	for idx in range(0, len(listFilePath)): listMethod.append("")

	if not isinstance(flg_issorted, bool): flg_issorted = False
	flg_issorted = not flg_issorted

	#*** Operations ********************************************************************************
	#--- Sort --------------------------------------------------------------------------------------
	if flg_issorted: listFilePath = hs_dataset.IBase_get_sorted_list([listFilePath], [0], [1])[0]

	#--- Content Analysis --------------------------------------------------------------------------
	flg_do   	= True
	cnt_main 	= 0
	wrkFilename = ""

	while flg_do:
		if listMethod[cnt_main] == "":
			curPath  	= listFilePath[cnt_main]
			curFilename = IBase_get_splitted_strpath(curPath)[2]

			if curFilename != wrkFilename: cnt_method = 0
			readBase    = IBase_read_binary(curPath)
			cnt_method  = cnt_method + 1
			listMethod[cnt_main] = cnt_method

			for idx in range(cnt_main, len(listFilePath)):
				compPath 	 = listFilePath[idx]
				compFilename = IBase_get_splitted_strpath(compPath)[2]
				
				if compFilename != curFilename: break
				elif listMethod[idx] == "":
					readComp = IBase_read_binary(compPath)
					flg_same = IBase_compare_list_issame(readBase, readComp)
					if flg_same: listMethod[idx] = cnt_method

			wrkFilename = curFilename

		cnt_main = cnt_main + 1
		if cnt_main >= len(listFilePath): flg_do = False

	#--- Release -----------------------------------------------------------------------------------
	return [listFilePath, listMethod]

def IUser_get_valid_name_for_creation(pathDest, nameFile, filetype, flg_tstamp):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a valid name for folder/file creation

	[str]  pathDest, 	  Target file's destination, Default is desktop
	[str]  nameFile, 	  Target name
	[str]  filetype, 	  File extension
	[bool] flg_tstamp,    True: Add timestamp to 'nameFolder' first before adding version

	'''

	#*** Input Validation **************************************************************************
	if not os.path.exists(pathDest): return 101
	if not (isinstance(nameFile, str) and len(nameFile) > 0): return nameFile

	#*** Initialization ****************************************************************************
	if not isinstance(flg_tstamp, bool): flg_tstamp = True

	chr_path  	 = getconst_chr_path()[1]
	strTimestamp = IBase_get_timestamp("yyyymmdd_hhmmss")
	nameUsed 	 = nameFile

	#*** Operations ********************************************************************************
	while os.path.exists(pathDest + chr_path + nameUsed + filetype):
		if flg_tstamp:
			nameUsed   = nameFile + "_" + strTimestamp
			flg_tstamp = False
		else:
			cnt = cnt + 1
			nameUsed = nameFile + "_" + str(cnt)
	return nameUsed

def IUser_create_folder(strRoot, nameFolder, flg_tryremove, flg_tstamp):
	#*** Documentation *****************************************************************************
	'''Documentation

		create a folder with a name as defined by nameFolder under strRoot directory
		if nameFolder already exist, the counter tag is added automatically to the folder name

	[str]  strRoot, 	  Root directory where nameFolder is expected to be created under
						  If strRoot doesn't exist, it will be created automatically.
	[str]  nameFolder, 	  Folder name, can't be empty
	[bool] flg_tryremove, True: Try to remove existing folder first before adding version 
	[bool] flg_tstamp,    True: Add timestamp to 'nameFolder' first before adding version

	'''

	#*** Input Validation **************************************************************************
	strRoot    = str(strRoot)
	nameFolder = str(nameFolder)

	if not len(strRoot) > 0: return 101

	#*** Initialization ****************************************************************************
	if not (isinstance(nameFolder, str) and len(nameFolder) > 0):
		tmpVal     = os.path.split(strRoot)
		strRoot    = tmpVal[0]
		nameFolder = tmpVal[1]

	if not isinstance(flg_tryremove, bool): flg_tryremove = False
	if not isinstance(flg_tstamp, bool): 	flg_tstamp    = True

	chr_path = getconst_chr_path()[1]
	strRoot  = IBase_get_formatted_path(strRoot, chr_path)
	nameOrg  = nameFolder
	cnt  	 = 0
	
	#*** Operations ********************************************************************************
	#--- Verify Root directory ---------------------------------------------------------------------
	if os.path.isfile(strRoot): 	 return 210
	if not os.path.exists(strRoot): 
		try: 	os.makedirs(strRoot)
		except: return 211

	#--- Check Existence of nameFolder under Root directory ----------------------------------------
	while os.path.exists(strRoot + chr_path + nameFolder):
		if flg_tryremove:
			resRemove = IUser_delete_content(strRoot, nameFolder)
			flg_tryremove = False
		else: resRemove = False

		if not resRemove:
			if flg_tstamp:
				nameOrg    = nameOrg + "_" + IBase_get_timestamp("yyyymmdd_hhmmss")
				nameFolder = nameOrg
				flg_tstamp = False
			else:
				cnt = cnt + 1
				nameFolder = nameOrg + "_" + str(cnt)

	#--- Create folder ----------------------------------------------------------------------------- 
	try: os.makedirs(strRoot + chr_path + nameFolder)
	except: return 230

	#--- Release -----------------------------------------------------------------------------------
	return nameFolder

def IUser_delete_content(strRoot, nameContent):
	#*** Documentation *****************************************************************************
	'''Documentation

		Delete target file/folder

	[str]  strRoot, 	Root directory where nameContent does exist
	[str]  nameContent, File or Folder name, can be empty

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(strRoot, str) and len(strRoot) > 0): return 101

	#*** Initialization ****************************************************************************
	chr_path = getconst_chr_path()[1]
	strRoot  = IBase_get_formatted_path(strRoot, chr_path) + chr_path

	#*** Operation *********************************************************************************
	#--- Take last folder instead if 'nameContent' is empty ----------------------------------------
	if not (isinstance(nameContent, str) and len(nameContent)):
		tmpInfo  	= os.path.split(strRoot)
		strRoot  	= tmpInfo[0]
		nameContent = tmpInfo[1]

	#--- Attemp to delete --------------------------------------------------------------------------
	if os.path.exists(strRoot + nameContent):
		try: 		os.remove(strRoot + nameContent) 		# Assume Content is a file
		except:
			try: 	shutil.rmtree(strRoot + nameContent) 	# Assume Content is a directory
			except: return False

	#--- Release -----------------------------------------------------------------------------------
	return True

def IUser_transfer_content(strSource, strDest, strNewName, flg_copy, flg_mirror):
	#*** Documentation *****************************************************************************
	'''Documentation

		Move/Copy target file to the target destination
		or
		Mirrors source content to target destination (only work if strSource is directory)

	[str]  strSource, 	Path of a target file or folder to be moved
	[str]  strDest, 	Destination folder path, If it doesn't exist, it will be created automatically
	[str]  strNewName,  New file name, provides "" if Renaming is not necessary
	[bool] flg_copy, 	Operation to be performed, True: Copy target File, False: Cut target File
	[bool] flg_mirror,  True: Mirror mode, strDest has same content as strSource, False: Normal
						flg_mirror will override flg_copy option
	
	'''

	#*** Input Validation **************************************************************************
	strSource = IBase_get_formatted_path(str(strSource), "/")
	strDest   = IBase_get_formatted_path(str(strDest), "/")
	
	if len(strSource) == 0 or len(strDest) == 0: 	return 101
	if not os.path.exists(strSource): 				return 102

	#*** Initialization ****************************************************************************
	if not (isinstance(flg_copy, bool)): 	flg_copy = False
	if not (isinstance(flg_mirror, bool)): 	flg_mirror = False

	strNewName   = str(strNewName)
	chr_quote    = getconst_chr_quote()
	chr_path 	 = getconst_chr_path()[1]
	strRoot   	 = ""
	sourceFolder = ""
	sourceFile   = ""

	#*** Operations ********************************************************************************
	#--- Source and Destination Preparation --------------------------------------------------------
	if os.path.isfile(strSource):
		strRoot   	 = os.path.split(strSource)[0]
		sourceFile   = os.path.split(strSource)[1]
		sourceFolder = os.path.split(strRoot)[1]

	elif os.path.isdir(strSource):
		strRoot   	 = strSource
		sourceFolder = os.path.split(strSource)[1]

	strRoot = IBase_get_formatted_path(strRoot, chr_path)
	strDest = IBase_get_formatted_path(strDest, chr_path)
	if sourceFile == "" and not flg_mirror:
		if len(strNewName) > 0: strDest = strDest + chr_path + strNewName
		else: 					strDest = strDest + chr_path + sourceFolder

	#--- Command Preparation -----------------------------------------------------------------------
	cmd_str	= "robocopy " + chr_quote + strRoot + chr_quote + " " + chr_quote + strDest + chr_quote
	
	if sourceFile != "": cmd_str = cmd_str + " " + chr_quote + sourceFile + chr_quote
	elif flg_mirror:	 cmd_str = cmd_str + " /mir"
	else: 				 cmd_str = cmd_str + " /e"

	if not flg_copy: 	 cmd_str = cmd_str + " /MOVE"

	#--- Deploy Content ----------------------------------------------------------------------------
	RetVal = os.system("cmd /c " + chr_quote + cmd_str + chr_quote)

	#--- Renaming: Extension for File only ---------------------------------------------------------
	if sourceFile != "" and len(strNewName) > 0:
		if sourceFile.rfind(".") > 0: FileType = sourceFile[sourceFile.rfind("."):]
		else: FileType = ""

		# Prevent double FileExtension
		if strNewName.find(FileType) != -1: FileType = ""

		try: os.rename(strDest + chr_path+ sourceFile, strDest + chr_path + strNewName + FileType)
		except: return True

	#--- Release -----------------------------------------------------------------------------------
	return True

def IUser_create_txt_file_fromlist(nameFile, path_dest, listInfo, chr_separator):
	#*** Documentation *****************************************************************************
	'''Documentation

		Create .txt file named 'nameFile' at 'path_dest' and write contents from 'listInfo' to it
		If 'nameFile' is not provided or is invalid then "PyGem_txt_" + timestamp will be used.

		If 'listInfo' has no list inside then each element in 'listInfo' will be separated by a
		linebreaker

		If 'listInfo' has lists inside then each element on each column will be separated by
		'chr_separator' and each column will be separated by 'linebreaker'

		If 'chr_separator' is not provided or is invalid then ";" is used.

	[str]  nameFile, 	  A name of this .txt file
	[str]  path_dest,  	  A string path where this .txt file will be created at
	[list] listInfo, 	  A list of content to be read into .txt file
	[str]  chr_separator, A character used as a separator between each element on each column in
							'listInfo'
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(path_dest, str) and os.path.exists(path_dest)): return False

	#*** Initialization ****************************************************************************
	if not (isinstance(chr_separator, str) and len(chr_separator) > 0): chr_separator = ";"
	if not (isinstance(nameFile, str) and len(nameFile) > 0): 
		nameFile = "PyGem_txt_" + IBase_get_timestamp("yyyymmdd_hhmmss")
	
	chr_path   = getconst_chr_path()[1]
	chr_lb 	   = "\n"
	filetype   = ".txt"
	listUsed   = []
	flg_single = False

	#*** Operations ********************************************************************************
	#--- Conditioning of 'listInfo' ----------------------------------------------------------------
	if sum([isinstance(x, list) for x in listInfo]) == 0: flg_single = True

	if flg_single: listUsed = list(listInfo)
	else: listUsed = [x if isinstance(x, list) else [x] for x in listInfo]

	#--- Create and Write txt file -----------------------------------------------------------------
	path_full = chr_path.join([path_dest, nameFile + filetype])

	with open(path_full, "w") as datfile:
		if flg_single: datfile.write(chr_lb.join(listUsed))
		else:
			for eachCol in listUsed:
				strComb = chr_separator.join(eachCol)
				datfile.write(strComb + chr_lb)
		datfile.close()

	#--- Release -----------------------------------------------------------------------------------
	return path_full

def IUser_create_file_fromstr(nameFile, pathDest, strInfo, fileType, flg_tryremove, flg_tstamp):
	#*** Documentation *****************************************************************************
	'''Documentation

		Create file that is written 'strInfo' content into it.

	[str]  nameFile, 	  Target filename, Default is "bbs_tmpfile_" + timestamp
	[str]  pathDest, 	  Target file's destination, Default is desktop
	[str]  strInfo, 	  Content to be written in this file
	[str]  fileType,      File extension
	[bool] flg_tryremove, True: Try to remove existing folder first before adding version 
	[bool] flg_tstamp,    True: Add timestamp to 'nameFolder' first before adding version

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(strInfo, str): return False

	#*** Initialization ****************************************************************************
	strTimestamp = IBase_get_timestamp("yyyymmdd_hhmmss")
	defFilename  = "bbs_tmpfile_" + strTimestamp
	defFiletype  = ".txt"
	defPath  	 = IBase_get_desktop_path()

	if not isinstance(flg_tryremove, bool): flg_tryremove = False
	if not isinstance(flg_tstamp, bool): 	flg_tstamp    = True

	chr_path = getconst_chr_path()[1]
	pathDest = IBase_get_formatted_path(pathDest, chr_path)
	cnt  	 = 0
	
	#*** Operations ********************************************************************************
	#--- Destination and File parameters preparation -----------------------------------------------
	if not (isinstance(pathDest, str) and len(pathDest) > 0): pathDest = defPath
	if not (isinstance(nameFile, str) and len(nameFile) > 0): nameFile = defFilename
	if not (isinstance(fileType, str) and len(fileType) > 0): fileType = defFiletype

	#--- Check Existence of 'nameFile' at 'pathDest' -----------------------------------------------
	nameUsed = nameFile

	while os.path.exists(pathDest + chr_path + nameUsed + fileType):
		if flg_tryremove:
			resRemove = IUser_delete_content(pathDest, nameUsed + fileType)
			flg_tryremove = False
		else: resRemove = False

		if not resRemove:
			if flg_tstamp:
				nameUsed   = nameUsed + "_" + strTimestamp
				flg_tstamp = False
			else:
				cnt = cnt + 1
				nameUsed = nameFile + "_" + str(cnt)

	#--- Create file -------------------------------------------------------------------------------
	with open(pathDest + chr_path + nameUsed + fileType, 'w') as outfile:
		outfile.write(strInfo)
		outfile.close()

    #--- Release -----------------------------------------------------------------------------------
	return pathDest + chr_path + nameUsed + fileType



#*** Function Group: Text file parser **************************************************************
def IUser_get_tag_value_from_text(pathFile, listTag, chr_merger, chr_separator):
	#*** Documentation *****************************************************************************
	'''Documentation

		Return a list of value found on each tag provided in 'listTag'
		Value is merged with each tag by 'chr_merger'.
		If 'chr_merger' is not provided, ":" is used by default.

		e.g. "TestCellId": "HDE1"	-> tag = "TestcellId", chr_merger = ":"
			 "driftMethod"= "2" 	-> tag = "driftMethod", chr_merger = "="

		Result Format
		[[tag1, [value1, value2, ...]], [tag2, [value1, value2, ....]], ...]

	[list] pathFile,      A list or string of file's path to be parsed
	[list] listTag,       A list or string of target tag that their value shall be collected
	[str]  chr_merger,    A character used as a merger between each tag and each value
	[str]  chr_separator, A character used as a separator between each tag group. If nothing is
						  provided then linebreaker will be used as default

	'''

	#*** Input Validation **************************************************************************
	if not len(pathFile) > 0: return 101
	if not len(listTag) > 0:  return 102

	#*** Initialization ****************************************************************************
	if not (isinstance(chr_merger, str) and len(chr_merger) > 0): 		chr_merger 	  = ":"
	if not (isinstance(chr_separator, str) and len(chr_separator) > 0): chr_separator = "\n"

	listTag  = hs_prep_StrList(listTag)
	pathFile = hs_prep_StrList(pathFile)
	chr_quot = chr(34)
	listRes  = []

	#*** Operations ********************************************************************************
	for eachFile in pathFile:
		if os.path.isfile(eachFile):
			with open(eachFile, "r") as tmpfile: strContent = tmpfile.read()

			for curTag in listTag:
				tmpRes  = [curTag, []]
				idx_max = len(strContent)
				idx_sta = 0

				while idx_sta <= idx_max:
					idx_sta = strContent.find(curTag, idx_sta)

					if idx_sta == -1: idx_sta = idx_max + 1
					else:
						idx_sta = strContent.find(chr_merger, idx_sta) + 1
						
						while strContent[idx_sta] == " ": idx_sta = idx_sta + 1

						if strContent[idx_sta] == chr_quot:
							idx_sta = idx_sta + 1
							idx_end = strContent.find(chr_quot, idx_sta)
						else: idx_end = strContent.find(chr_separator, idx_sta)

						if idx_end == -1: idx_sta = idx_sta + 1 	# Something wrong, skip to next
						else:
							curVal  = strContent[idx_sta:idx_end]
							tmpRes[1].append(curVal)
							idx_sta = idx_end + 1

				listRes.append(tmpRes)

	#--- Release -----------------------------------------------------------------------------------
	return listRes


