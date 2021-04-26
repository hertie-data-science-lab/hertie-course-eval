import streamlit as st
import pandas as pd
import pathlib
from zipfile import ZipFile
import glob, os
from io import StringIO
import os
import base64

# Note for future dev:
# Raw course files are in .xls extension and should be cleaned out after processing
# I don't refractor because I'm an unorganized masochist, please don't judge


def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download Processed {file_label} File</a>'
    return href

def clean_folder(file_extension):
	directory = os.getcwd()

	files_in_directory = os.listdir(directory)
	filtered_files = [file for file in files_in_directory if file.endswith(file_extension)]
	for file in filtered_files:
		path_to_file = os.path.join(directory, file)
		os.remove(path_to_file)

def main():

# local_css("style.css")

	st.image("lab.png")

	st.title('Hertie School Course Evaluation App')

	activities = ["Consolidated Report MPP/MIA", "Individual Report MPP/MIA"]
	choice = st.sidebar.selectbox("Pick reporting type", activities)

	if choice == "Consolidated Report MPP/MIA":

		st.write("""
				The app below helps to automate the process of creating consolidated data for the course evaluation at the Hertie School. \n
				To start, please select the raw files of all courses that you would like to evaluate together and make one compressed zip file containing all of them.
				You can upload the zip file below, the app will automatically process all files and create a consolidated data table, which you can download to be used for the report.\n 

				The app assumes that all the main questions are placed in the sheet named "1-10" in the raw files.
			""")

		# You can specify more file types below if you want
		zip_file = st.file_uploader("Upload zip file for courses to be evaluated together", type=['zip'])

		if zip_file is not None:
			# Extract the zip file
			ZipFile(zip_file).extractall()

			# Get the number of excel files in the current directory
			myPath = pathlib.Path().absolute()
			fileCounter = len(glob.glob1(myPath,"*.xls"))

			# Get a list of files in xls
			filenames = glob.glob("*.xls")

			# get a list of dataframes
			list_of_sheets = [pd.read_excel(filename, sheet_name=None) for filename in filenames]

			df = []
			new_list_of_df = []

			for sheet in list_of_sheets: 

				df = sheet['1-10']
				new_df = df[['Question', 'Total', 'Average', 'Median', 'Variance']].copy()
				new_df["Average"] = new_df["Average"].str.replace(",",".").astype(float)
				new_df["Median"] = new_df["Median"].str.replace(",",".").astype(float)
				new_df["Variance"] = new_df["Variance"].str.replace(",",".").astype(float)

				df2 = sheet['0-100']
				df2 = df2[['Question', 'Total', 'Average', 'Median', 'Variance']].copy()
				df2["Average"] = df2["Average"].str.replace(",",".").astype(float)
				df2["Median"] = df2["Median"].str.replace(",",".").astype(float)
				df2["Variance"] = df2["Variance"].str.replace(",",".").astype(float)
				df2['Question'] = "Weighted Average all core courses (Grading scale 0-100)"
				df2 = df2.head(1)
				

				# If there are more questions than 13, it means there are 2 professors, clean up the dataframe by combining the 2 professors' evaluation
				if len(new_df.index) > 14: 
					academic = new_df.loc[new_df['Question'].str.startswith('Provision of academic insights')]
					practical = new_df.loc[new_df['Question'].str.startswith('Provision of practical insights')]
					expertise = new_df.loc[new_df['Question'].str.startswith('Expertise of instructor')]
					availability = new_df.loc[new_df['Question'].str.startswith('Availability of instructor outside of class')]
					responsiveness = new_df.loc[new_df['Question'].str.startswith('Responsiveness of instructor in class')]
					fairness = new_df.loc[new_df['Question'].str.startswith('Fairness of feedback to assignments')]
					timeliness = new_df.loc[new_df['Question'].str.startswith('Timeliness of feedback to assignments')]

					academic["Average"] = academic["Average"].mean()
					academic["Median"] = academic["Median"].mean()
					academic["Variance"] = academic["Variance"].mean()
					academic['Question']= "Provision of academic insights (10=highly satisfied)"
					academic = academic.head(1)

					practical["Average"] = practical["Average"].mean()
					practical["Median"] = practical["Median"].mean()
					practical["Variance"] = practical["Variance"].mean()
					practical['Question']= "Provision of practical insights (10=highly satisfied)"
					practical = practical.head(1)

					expertise["Average"] = expertise["Average"].mean()
					expertise["Median"] = expertise["Median"].mean()
					expertise["Variance"] = expertise["Variance"].mean()
					expertise['Question']= "Expertise of instructor (10=highly satisfied)"
					expertise = expertise.head(1)

					availability["Average"] = availability["Average"].mean()
					availability["Median"] = availability["Median"].mean()
					availability["Variance"] = availability["Variance"].mean()
					availability['Question']= "Availability of instructor outside of class (10=highly satisfied)"
					availability = availability.head(1)

					responsiveness["Average"] = responsiveness["Average"].mean()
					responsiveness["Median"] = responsiveness["Median"].mean()
					responsiveness["Variance"] = responsiveness["Variance"].mean()
					responsiveness['Question']= "Responsiveness of instructor in class (10=highly satisfied)"
					responsiveness = responsiveness.head(1)

					fairness["Average"] = fairness["Average"].mean()
					fairness["Median"] = fairness["Median"].mean()
					fairness["Variance"] = fairness["Variance"].mean()
					fairness['Question']= "Fairness of feedback to assignments (10=highly satisfied)"
					fairness = fairness.head(1)

					timeliness["Average"] = timeliness["Average"].mean()
					timeliness["Median"] = timeliness["Median"].mean()
					timeliness["Variance"] = timeliness["Variance"].mean()
					timeliness['Question']= "Timeliness of feedback to assignments (10=highly satisfied)"
					timeliness = timeliness.head(1)

					# Delete old dataframes with 2 instructors

					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Provision of academic insights')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Provision of practical insights')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Expertise of instructor')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Availability of instructor outside of class')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Responsiveness of instructor in class')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Fairness of feedback to assignments')]
					new_df = new_df[~new_df['Question'].astype(str).str.startswith('Timeliness of feedback to assignments')]

					# Weighted average total



					# Add final dataframe for consolidated file

					frames = [new_df, academic, practical, expertise, availability, responsiveness, fairness, timeliness]
					new_df = pd.concat(frames)
					new_list_of_df.append(new_df)
				else:
					new_list_of_df.append(new_df)


				# Calculate final consolidated table
				final_df = new_list_of_df[0]

				final_df['Total Response Rate'] = 0
				final_df['Average Median all courses'] = 0
				final_df['Average Variance all courses'] = 0
				final_df['Average all courses not weighted'] = 0
				final_df['Weighted average all courses'] = 0

				for df in new_list_of_df: 
					final_df['Total Response Rate'] += df['Total']
					final_df['Average Median all courses'] += df['Median'] / fileCounter
					final_df['Average Variance all courses'] += df['Variance'] / fileCounter
					final_df['Average all courses not weighted'] +=  df['Total'] *  df['Average']

				final_df['Weighted average all courses'] = final_df['Average all courses not weighted'] / final_df['Total Response Rate']
				final_df = final_df.drop(['Total', 'Average', 'Median', 'Variance', 'Average all courses not weighted'], axis = 1)

				# Round it to 2 decimals place
				final_df = final_df.round(decimals=2)

			st.dataframe(final_df)
			final_df.to_excel("Consolidated_Data.xlsx") 

			st.markdown(get_binary_file_downloader_html('Consolidated_Data.xlsx', 'Excel'), unsafe_allow_html=True)

			clean_folder(".xls")
			clean_folder(".zip")

	elif choice == "Individual Report MPP/MIA":

		st.write("""

				The app below helps to automate the process of creating individual reporting data for the course evaluation at the Hertie School. \n
				To start, please select the raw files of all courses that you would like to evaluate and make one commpressed zip file containing all of them.
				You can upload the zip file below, new individual excel files will be generated in a new zipped file that you can download.\n 

			""")

		# You can specify more file types below if you want
		zip_file = st.file_uploader("Upload zip file for all courses to be evaluated", type=['zip'])

		if zip_file is not None:
			# Extract the zip file
			ZipFile(zip_file).extractall()

			# Get the number of excel files in the current directory
			myPath = pathlib.Path().absolute()
			fileCounter = len(glob.glob1(myPath,"*.xls"))

			# Get a list of files in xls
			filenames = glob.glob("*.xls")

			# get a list of dataframes
			list_of_sheets = [pd.read_excel(filename, sheet_name=None) for filename in filenames]

			df = []
			new_list_of_df = []
			file_counter = 0 #set counter for the excel file

			# create a ZipFile object
			zipObj = ZipFile('Individual_reports.zip', 'w')

			for sheets in list_of_sheets: 
			    df = sheets["Fulltext Questions"]    
			    df2 = sheets["0-100"]
			    df2 = df2.loc[:, ~(df2 == 0).any()]
			    df3 =sheets["1-3"]
			    df4 = sheets["1-4"]
			    df5 = sheets["1-10"]
			    
			    # Process question tab
			    df = df[df.Answer != "-"]
			    df = df[df.Answer != "/"]
			    df = df[df.Answer != "."]
			    df = df[df.Answer != "No comment"]
			    df = df[df.Answer != "no comment"]
			    df = df[df.Answer != "n/a"]
			    df = df[df.Answer != "na"]
			    df = df[df.Answer != "N/A"]
			    df = df[df.Answer != "NA"]
			    df = df[df.Answer != "none"]
			    df = df[df.Answer != "None"]
			    df = df[df.Answer != "Not applicable"]
			    df = df[df.Answer != "not applicable"]
			    df = df[df.Answer != "no guests"]
			    df = df[df.Answer != "no guest"]
			    df = df[df.Answer != "No guest"]
			    df = df[df.Answer != "No guests"]
			    df = df[df.Answer.notnull()]

			    # Put multiple dataframes into one xlsx sheet

			    # funtion
			    def multiple_dfs(df_list, sheets, file_name, spaces):
			        writer = pd.ExcelWriter(file_name)   
			        row = 0
			        for dataframe in df_list:
			            dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)   
			            row = row + len(dataframe.index) + spaces + 1
			        writer.save()

			    # list of dataframes
			    dfs = [df2, df3, df4, df5]

			    # # run function

			    multiple_dfs(dfs, '1-10', f'{filenames[file_counter]}.xlsx', 1)

			    from openpyxl import load_workbook

			    excel_path = f'{filenames[file_counter]}.xlsx'

			    book = load_workbook(excel_path)
			    with pd.ExcelWriter(excel_path) as writer:
			        writer.book = book
			        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)    

			        ## Your dataframe to append. 
			        df.to_excel(writer, 'Fulltext Questions')  

			        writer.save()
			        

			    # Add multiple files to the zip
			    zipObj.write(f'{filenames[file_counter]}.xlsx')

			    file_counter += 1 # increment to get name of the next excel file
			    
			zipObj.close()

			st.markdown(get_binary_file_downloader_html('Individual_reports.zip', 'Zip'), unsafe_allow_html=True)

			clean_folder(".xls")


if __name__ == "__main__":
    main()
