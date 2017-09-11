import time
import os
from os.path import isfile, join
from datetime import datetime
from Excel_Comparison import Excel_Comparison

def get_downloaded_file(dic_path):
	#get files list from path
	files = [f for f in os.listdir(dic_path) if isfile(join(dic_path, f))]
	#check file name with SearchResult in the beginning
	search_files = [ [f, os.stat(join(dic_path, f)).st_mtime] for f in files if f.startswith("SearchResults") ]
	# search_files = [ f for f in files if f.startswith("SearchResults") ]
	search_files = sorted(search_files, key=lambda x : x[1], reverse=True)
	return [i[0] for i in search_files]

def get_issue_lists(dic_path, proj_name):
	#get all files in dic
	files = [f for f in os.listdir(dic_path) if isfile(join(dic_path, f))]
	#get any files that has proj_name in the begining of the file name
	search_files = [join(dic_path, f) for f in files if f.startswith(proj_name)]
	#sort by file name (newest in the front)
	search_files.sort(reverse=True)
	#return the list
	return search_files


def rename_issue_list(proj_name, download_path, store_path):
	files = get_downloaded_file(download_path)
	if len(files) > 0:
		#create rename title
		time_stamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
		original_name = join(download_path,files[0])
		modified_name = join(store_path,proj_name+"_issue_"+time_stamp+".xls")
		print original_name, modified_name
		#rename search result
		os.rename(original_name, modified_name)

def compare_excels(proj_name, path):
	excel_files = get_issue_lists(path, proj_name)
	if len(excel_files) > 2:
		 excel = Excel_Comparison(excel_files[0], excel_files[1])
		 excel.compare()
	else:
		logger.info("only 1 issue list in folder")