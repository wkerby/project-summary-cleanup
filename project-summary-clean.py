# #import pandas for excel data
import pandas as pd

#create dataframes
proj_df = pd.read_excel(r"Z:\\Users\\WKerby\\My Computer\\Documents\\KMeeks\\Project Summary.xlsx")
copy_df = pd.read_excel(r"Z:\\Users\\WKerby\\My Computer\\Documents\\KMeeks\\Copy of Project Summary.xlsx")
copy_dct = copy_df.to_dict()


#establish list of indices for the Copy of Project Summary Sheet
pindices = proj_df.index
cindices = copy_df.index

#establish list of column headers for Copy of Project Summary Sheet
copy_headers = list(copy_df)

#create index list for Project Summary and Copy of Project Summary
job_indices = list(proj_df.loc[:, "Job"])
copy_job_indices = list(copy_df.loc[:, "Job"])

#create df from which the cleaned Project Summary sheet will be created
copy_copy_dict = {}
for header in copy_headers:
	copy_copy_dict[header] = {}

for index in list(cindices):
	for header in copy_headers:
		copy_copy_dict[header][index] = ""

#loop through job numbers from Copy of Project Summary Sheet, and if job number appears in Project Summary sheet, fill each 
#cell of data from Project Summary Sheet into row in the Copy of Project Summary Sheet
for job in copy_job_indices:
	if job in job_indices:
		cindex = cindices[copy_df["Job"] == job]
		pindex = pindices[proj_df["Job"] == job]
		for header in copy_headers:
			copy_copy_dict[header][cindex[0]] = proj_df.loc[pindex,header]
			# copy_df.loc[cindex,header] = proj_df.loc[pindex,header]

print(copy_copy_dict)

#write to excel
copy_copy_df = pd.DataFrame(copy_copy_dict)
try:
	copy_copy_df.to_excel(r"Z:\\Users\\WKerby\\My Computer\\Documents\\KMeeks\\Copy of Project Summary_.xlsx", index = False)
except:
	print("failed")
else:
	print("succeeded")


