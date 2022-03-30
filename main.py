import pandas as pd
import numpy as np
import win32com.client as win32
import os
import xlsxwriter

def getFiles(directory):
    '''given a directory, will return list of all files in given directory'''
    fileList = []  
    if directory[-1] != '\\': 
        directory = directory + '\\'  
    for item in os.scandir(directory):
        if item.is_file():  
            fileList.append(directory + item.name)  
    return fileList  

def convertFiles(fileList):
    '''given a list of file paths, will save .xls, .xlsx, .xlsm, or .csv files as .xlsb in same directory'''
    for file in fileList: 
        if os.path.splitext(file)[1] in ['.xls', '.xlsx', '.xlsm', '.csv']: 
            tgtPath = os.path.splitext(file)[0] + '.xlsb'  
            xlApp = win32.Dispatch('Excel.Application')  
            xlApp.Visible = False 
            xlApp.ScreenUpdating = False  
            xlApp.DisplayAlerts = False  
            try:
                wb = xlApp.Workbooks.Open(Filename=file, ReadOnly=True)  
            except:
                print(f'Could not open {file}')  
                continue
            wb.SaveAs(Filename=tgtPath, FileFormat=50)
            wb.Close(False)  
            xlApp.Quit()  
            print(f'Saved {tgtPath} from {file}')  

def main(directory):
    '''given a windows formatted folder path, will convert all .xls, .xlsx, .xlsm, and .csv files in that directory
    to .xlsb format'''
    fileList = getFiles(directory)  # get list of files in provided directory
    convertFiles(fileList)  # convert Excel/CSV files in file list



base_file_ECN = pd.read_excel("fast_api/Test_Brt_Base_Mar22(2).xlsb",sheet_name="North_East_Central")
base_file_SW = pd.read_excel("fast_api/Test_Brt_Base_Mar22(2).xlsb",sheet_name="S1_S2_West")
basefile_combined = [base_file_ECN,base_file_SW]
base_file_combine = pd.concat(basefile_combined)


tracker_dump_ECN_verified = (base_file_ECN.loc[base_file_ECN['Status'] == "Verified"]) 
ECN_retag = (base_file_ECN.loc[base_file_ECN['Status'] == "Retag"])
tracker_dump_SW_verified = (base_file_SW.loc[base_file_SW['Status'] == "Verified"]) 
SW_retag = (base_file_SW.loc[base_file_SW['Status'] == "Retag"])
retag_combined = [ECN_retag,SW_retag]
retag_file_combine = pd.concat(retag_combined)


file_name_ECN_verified = "C:\\Users\\91859\\Fast_docs\\dump(East_Central_North).xlsx"
file_name_SW_verified = "C:\\Users\\91859\\Fast_docs\\dump(S1_S2_West).xlsx"
#file_name_ECN_retag = "C:\\Users\\91859\\Fast_docs\\retag(East_Central_North).xlsx"
#file_name_SW_retag = "C:\\Users\\91859\\Fast_docs\\retag(S1_S2_West).xlsx"
file_name_retag = "C:\\Users\\91859\\Fast_docs\\retag.xlsx"
"""
tracker_dump_ECN_verified.to_excel(file_name_ECN_verified,index=False)
#ECN_retag.to_excel(file_name_ECN_retag,index=False)
tracker_dump_SW_verified.to_excel(file_name_SW_verified,index=False)
#SW_retag.to_excel(file_name_SW_retag,index=False)"""
retag_file_combine.to_excel(file_name_retag,index=False)

#list of top 200 AW_Code.
t_200 = [13, 29, 194, 756, 1434, 1449, 1780, 1786, 1806, 2006, 2071, 2205, 2209, 2218, 2290, 2407, 2729, 2888, 
4355, 4912, 5190, 7359, 7865, 8505, 8572, 9887, 9893, 10317, 10369, 11222, 11494, 11812, 12120, 12907, 
13056, 13058, 13953, 14129, 14364, 14398, 15349, 15551, 15680, 15773, 15808, 15915, 16222, 16288, 16868, 17449, 17467, 17788, 18365, 18694, 18786, 18873, 18928, 19073, 19436, 19535, 19544, 19690, 20164, 20243, 20568, 20680, 20894, 21009, 21011, 21223, 21997, 22112, 22140, 22188, 22269, 22911, 23147, 23232, 23529,
23846, 23871, 24056, 24089, 24177, 24241, 24406, 24572, 24616, 24646, 24648, 24985, 25002, 25029, 25054, 25068, 25102, 25139, 25162, 25311, 25455, 25505, 25518, 25528, 25584, 25589, 25593, 25663, 25693, 25748, 25753, 25755, 25761, 25765, 25821, 25972, 25980, 25985, 26049, 26125, 26459, 26494, 26523, 26541, 26644, 26688, 26781, 26796, 26797, 26937, 27087, 27168, 27173, 27187, 27237, 27246, 27250, 27254, 
27309, 27352, 27436, 27581, 27661, 27688, 27763, 27790, 27855, 27859, 27923, 28000, 28042, 28047, 28050, 28086, 28104, 28143, 28374, 28396, 28402, 28434, 28444, 28490, 28594, 28622, 28658, 28730, 28806, 28829, 28839, 28849, 28957, 29150, 29288, 29337, 29338, 29348, 29360, 29486, 29512, 29524, 29621, 29727, 29774, 29826, 29909, 29926, 29929, 29953, 29962, 29984, 30011, 30114, 30122, 30201, 30255, 30267, 30277, 30303, 30319, 30322, 30329, 30343, 30402, 30440, 30508, 30515, 30534, 30543, 30546, 30558, 30718, 30776, 30890, 30994, 31055, 31062, 31098, 31172, 31193, 31200, 31206, 31210, 31232, 31234, 31236, 31297, 3131758, 31810, 31821, 31825, 31852, 31872, 31887, 958, 992, 1065, 1594, 1669, 2616, 2651, 3309, 3332,
7422, 8077, 8703, 9417, 10243, 10390, 10698, 11187, 11780, 12047, 12316, 12371, 13636, 13977, 14422, 15143, 15457, 15492, 15788, 16172, 16372, 16436, 16682, 16945, 17120, 18081, 18236, 18685, 18955, 19239, 19391, 19964, 20499, 21111, 21183, 21316, 21371, 21459, 22043, 22147, 22243, 22543, 22623, 22803, 22806, 
23094, 23531, 23717, 23910, 23942, 23943, 24159, 24161, 24262, 24286, 24291, 24444, 24539, 24812, 24840, 24898, 25003, 25109, 25152, 25174, 25210, 25226, 25351, 25358, 25396, 25572, 25575, 25727, 25885, 26101, 26103, 26104, 26402, 26768, 26799, 26827, 26836, 26947, 26965, 27210, 27283, 27337, 27347, 27450, 27515,
27687, 27706, 27767, 27826, 28040, 28096, 28102, 28147, 28253, 28332, 28445, 28456, 28464, 28465, 28731, 28746, 28784, 28790, 28909, 28934, 28936, 28943, 28987, 29035, 29068, 29088, 29093, 29098, 29169, 29270, 29398, 29419, 29440, 29509, 29515, 29582, 29677, 29744, 29753, 29810, 30243, 30248, 30386, 30393, 30400, 30576, 30612, 30760, 30839, 30905, 30913, 31135, 31153, 31180, 31213, 31230, 31340, 31389, 
31425, 31432, 31492, 31621, 31783, 31820, 31847, 31897]
t_200.sort()

AW_Summary_mix = base_file_combine.loc[:,'Region':'AW_Name']


AW_Summary_ECN = tracker_dump_ECN_verified.loc[:,'Region':'AW_Name']
AW_Summary_SW = tracker_dump_SW_verified.loc[:,'Region':'AW_Name']
combined_summary = [AW_Summary_ECN, AW_Summary_SW]  
AW_Summary___ = pd.concat(combined_summary)

not_in_dump = np.setdiff1d(AW_Summary_mix["AW_Code"], AW_Summary___["AW_Code"])
AW_Summary___ = AW_Summary___.append(pd.DataFrame({'AW_Code': not_in_dump}))
#AW_Summary___['AW_Code'] = AW_Summary___['AW_Code'].replace(['1'],'0')
# ^abhi add kra



ans_df = AW_Summary_mix.drop_duplicates(subset="AW_Code")
sorted_df = ans_df.sort_values(by='AW_Code')
#only_AW_column = sorted_df[['AW_Code']].copy() --ye use nhi hua
#unique_base_file = base_file_combine.drop_duplicates(subset="AW_Code")# ye use nhi hua


filtered_base_file = base_file_combine.loc[base_file_combine['AW_Code'].isin(sorted_df['AW_Code'])]
#unique_base_file = base_file_combine.drop_duplicates(subset="AW_Code") --ye use nhi hua
#sorted_df = sorted_df.fillna(0) ye use nhi hua
sorted_df.insert(5, "Geotagged Outlets",[x for x in (AW_Summary___.pivot_table(index = 'AW_Code', aggfunc ='size'))])
sorted_df.insert(6,"Base Outlets",[y for y in (filtered_base_file.pivot_table(index = 'AW_Code',aggfunc='size'))])
sorted_df['Geotagged Outlets'].mask(sorted_df['AW_Code'].isin(not_in_dump) ,0, inplace=True)

sorted_df1 = (sorted_df['Geotagged Outlets'].div(sorted_df['Base Outlets']).mul(100)).to_frame('Geotagged-base')
sorted_dff = sorted_df.copy()                      # Create copy of first DataFrame
sorted_dff["Geotagged-base"] = sorted_df1["Geotagged-base"] 
# adding birufication in AW summary
def f(sorted_dff):
    if sorted_dff['Geotagged-base'] < 50:
        val = '0-50%'
    elif sorted_dff['Geotagged-base'] > 50 and sorted_dff['Geotagged-base'] < 65:
        val = '50-65%'
    elif sorted_dff['Geotagged-base'] > 65 and sorted_dff['Geotagged-base'] < 80:
        val = '65-80%'
    else:
        val = '80-100%'
    return val
"""def g(sorted_dff):
    if sorted_dff['AW_Code'].isin(t_200):
        value = "Y"
    
    else:
        value = 'N'
    return value"""


sorted_dff['Bifurcation'] = sorted_dff.apply(f, axis=1)
#   sorted_dff['T 200'] = sorted_dff.apply(g, axis=1)
sorted_dff['T 200'] = np.where(sorted_dff['AW_Code'] in t_200, 'Y', 'N')


AW_Summary = sorted_dff
file_name_AW_summary = "C:\\Users\\91859\\Fast_docs\\AW_summary_output.xlsx"
AW_Summary.to_excel(file_name_AW_summary,index=False)

"""

# from here, dashboard starts
data = {'Region':  ['East', 'Central', 'North 1', 'North 2', 'South 1', 'South 2','West'],        
        }
Dashboard = pd.DataFrame(data)
ECN_retag = pd.read_excel("C:\\Users\\91859\\Fast_docs\\run_once\\retag(East_Central_North).xlsx")
SW_retag = pd.read_excel("C:\\Users\\91859\\Fast_docs\\run_once\\retag(S1_S2_West).xlsx")


#for AW code
e = AW_Summary['Region'].value_counts()['East']
c = AW_Summary['Region'].value_counts()['Central']
n1 = AW_Summary['Region'].value_counts()['North 1']
n2 = AW_Summary['Region'].value_counts()['North 2']
s1 = AW_Summary['Region'].value_counts()['South 1']
s2 = AW_Summary['Region'].value_counts()['South 2']
w = AW_Summary['Region'].value_counts()['West']



#for retag
only_east_column_re = ECN_retag['Region'].value_counts()['East']
only_central_column_re = ECN_retag['Region'].value_counts()['Central']
only_north1_column_re = ECN_retag['Region'].value_counts()['North 1']
only_north2_column_re = ECN_retag['Region'].value_counts()['North 2']
only_south1_column_re = SW_retag['Region'].value_counts()['South 1']
only_south2_column_re = SW_retag['Region'].value_counts()['South 2']
only_west_column_re = SW_retag['Region'].value_counts()['West']
list_retag_outlets = [only_east_column_re, only_central_column_re, only_north1_column_re, only_north2_column_re, only_south1_column_re, only_south2_column_re, only_west_column_re]




#for Geotag outlets
only_east_column_g = (AW_Summary.loc[AW_Summary['Region'] == "East"])["Geotagged Outlets"]
only_central_column_g = (AW_Summary.loc[AW_Summary['Region'] == "Central"])["Geotagged Outlets"]
only_north1_column_g = (AW_Summary.loc[AW_Summary['Region'] == "North 1"])["Geotagged Outlets"]
only_north2_column_g = (AW_Summary.loc[AW_Summary['Region'] == "North 2"])["Geotagged Outlets"]
only_south1_column_g = (AW_Summary.loc[AW_Summary['Region'] == "South 1"])["Geotagged Outlets"]
only_south2_column_g = (AW_Summary.loc[AW_Summary['Region'] == "South 2"])["Geotagged Outlets"]
only_west_column_g = (AW_Summary.loc[AW_Summary['Region'] == "West"])["Geotagged Outlets"]

geo_outlets_west = only_west_column_g.sum()
geo_outlets_east = only_east_column_g.sum()
geo_outlets_central = only_central_column_g.sum()
geo_outlets_north1 = only_north1_column_g.sum()
geo_outlets_north2 = only_north2_column_g.sum()
geo_outlets_south1 = only_south1_column_g.sum()
geo_outlets_south2 = only_south2_column_g.sum()
list_geo_outlets = [geo_outlets_east, geo_outlets_central, geo_outlets_north1, geo_outlets_north2, geo_outlets_south1, geo_outlets_south2, geo_outlets_west]




#for total outlets
only_east_column = (AW_Summary.loc[AW_Summary['Region'] == "East"])["Base Outlets"]
only_central_column = (AW_Summary.loc[AW_Summary['Region'] == "Central"])["Base Outlets"]
only_north1_column = (AW_Summary.loc[AW_Summary['Region'] == "North 1"])["Base Outlets"]
only_north2_column = (AW_Summary.loc[AW_Summary['Region'] == "North 2"])["Base Outlets"]
only_south1_column = (AW_Summary.loc[AW_Summary['Region'] == "South 1"])["Base Outlets"]
only_south2_column = (AW_Summary.loc[AW_Summary['Region'] == "South 2"])["Base Outlets"]
only_west_column = (AW_Summary.loc[AW_Summary['Region'] == "West"])["Base Outlets"]

total_outlets_west = only_west_column.sum()
total_outlets_east = only_east_column.sum()
total_outlets_central = only_central_column.sum()
total_outlets_north1 = only_north1_column.sum()
total_outlets_north2 = only_north2_column.sum()
total_outlets_south1 = only_south1_column.sum()
total_outlets_south2 = only_south2_column.sum()
list_total_outlets = [total_outlets_east, total_outlets_central, total_outlets_north1, total_outlets_north2, total_outlets_south1, total_outlets_south2, total_outlets_west]




Dashboard.insert(1, "AW Count",[e,c,n1,n2,s1,s2,w])
Dashboard.insert(2, "Total Outlets",list_total_outlets)
Dashboard.insert(3, "Geotagged Outlets",list_geo_outlets)
Dashboard.insert(4, "Retag Outlets",list_retag_outlets)
#print(Dashboard)
#Dashboard.to_excel("C:\\Users\\91859\\Fast_docs\\dashhboard.xlsx")


# from here Som based dashboard


# GEOTAGGED OUTLETS
data = {'SOM':  ['Abhijit Bera', 'Abhimanyu Singh', 'Abinash Patnaik', 'Amit Kumar Dhage', 'Atin Kapoor', 'Biswapriyo Roy','Abhishek saxena','GOKUL S','Philip P Koshy','Vikas Tandon','KASI VISHWANATHAN P','Biswarup Saha','Mahesh Gurumurthy','Manoj Razdan','Minesh Bhatt','Musir Rahman','Nihar Pandya','Nilesh Kumar','Nitesh Sood','Pardeep Pandey','Rajesh  Samantaray','RAMESH KUMAR A','Rejo Mathew','SAIKANNAN R','Sanjeev Kumar','Saravanan TN','Sudipta Mitra','Sumeet Nagpal','THANIS D','Mihir Kumar','Vishal Sharma','Lalit Mohan Joshi','Sanjeev Chopra','Shailendra Sharma','Shubham Kaushik',],}
Som_dataframee = pd.DataFrame({'SOM':  ['Abhijit Bera', 'Abhimanyu Singh', 'Abinash Patnaik', 'Amit Kumar Dhage', 'Atin Kapoor', 'Biswapriyo Roy','Abhishek saxena','GOKUL S','Philip P Koshy','Vikas Tandon','KASI VISHWANATHAN P','Biswarup Saha','Mahesh Gurumurthy','Manoj Razdan','Minesh Bhatt','Musir Rahman','Nihar Pandya','Nilesh Kumar','Nitesh Sood','Pardeep Pandey','Rajesh  Samantaray','RAMESH KUMAR A','Rejo Mathew','SAIKANNAN R','Sanjeev Kumar','Saravanan TN','Sudipta Mitra','Sumeet Nagpal','THANIS D','Mihir Kumar','Vishal Sharma','Lalit Mohan Joshi','Sanjeev Chopra','Shailendra Sharma','Shubham Kaushik',],}
)

#for geotag outlets
only_abhijit_bera_column_g = (AW_Summary.loc[AW_Summary['SOM'] == "Abhijit Bera"])["Geotagged Outlets"].sum()
only_abhimanyu_singh_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Abhimanyu Singh'])["Geotagged Outlets"].sum()
only_abinash_patnaik_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Abinash Patnaik'])["Geotagged Outlets"].sum()
only_amit_kumar_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Amit Kumar Dhage'])["Geotagged Outlets"].sum() 
only_atin_kapoor_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Atin Kapoor'])["Geotagged Outlets"].sum()

only_b_roy_column_g= (AW_Summary.loc[AW_Summary['SOM'] == 'Biswapriyo Roy'])["Geotagged Outlets"].sum()

only_a_saxena_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'ABHISHEK SAXENA'])["Geotagged Outlets"].sum()

only_gokul_s_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'GOKUL S'])["Geotagged Outlets"].sum()

only_philip_p_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Philip P Koshy'])["Geotagged Outlets"].sum()

only_v_tandon_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Vikas Tandon'])["Geotagged Outlets"].sum()

only_kasi_p_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'KASI VISHWANATHAN P'])["Geotagged Outlets"].sum()

only_b_saha_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Biswarup Saha'])["Geotagged Outlets"].sum()

only_m_gurumurthy_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Mahesh Gurumurthy'])["Geotagged Outlets"].sum()

only_manoj_razdan_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Manoj Razdan '])["Geotagged Outlets"].sum()

only_m_bhatt_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Minesh Bhatt'])["Geotagged Outlets"].sum()

only_m_rahman_column_g= (AW_Summary.loc[AW_Summary['SOM'] == 'Musir Rahman'])["Geotagged Outlets"].sum()

only_n_pandya_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Nihar Pandya'])["Geotagged Outlets"].sum()

only_n_kumar_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Nilesh Kumar'])["Geotagged Outlets"].sum()

only_n_sood_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Nitesh Sood'])["Geotagged Outlets"].sum()

only_p_pandey_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Pardeep Pandey'])["Geotagged Outlets"].sum()

only_r_samatarey_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Rajesh  Samantaray'])["Geotagged Outlets"].sum()

only_r_kumar_A_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'RAMESH KUMAR A'])["Geotagged Outlets"].sum()

only_r_mathew_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Rejo Mathew'])["Geotagged Outlets"].sum()

only_s_r_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'SAIKANNAN R'])["Geotagged Outlets"].sum()

only_sanjeev_k_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Sanjeev Kumar'])["Geotagged Outlets"].sum()

only_sarvanan_column_g =(AW_Summary.loc[AW_Summary['SOM'] == 'Saravanan TN'])["Geotagged Outlets"].sum()

only_s_mitra_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Sudipta Mitra'])["Geotagged Outlets"].sum()

only_s_nagpal_column_g =(AW_Summary.loc[AW_Summary['SOM'] == 'Sumeet Nagpal'])["Geotagged Outlets"].sum()

only_thanis_d_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'THANIS D'])["Geotagged Outlets"].sum()

only_m_kumar_column_g =(AW_Summary.loc[AW_Summary['SOM'] == 'Mihir Kumar'])["Geotagged Outlets"].sum()

only_v_sharma_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Vishal Sharma'])["Geotagged Outlets"].sum()

only_lalit_joshi_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Lalit Mohan Joshi'])["Geotagged Outlets"].sum()

only_s_chopra_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Sanjeev Chopra'])["Geotagged Outlets"].sum()

only_s_sharma_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Shailendra Sharma'])["Geotagged Outlets"].sum()
only_s_kaushik_column_g = (AW_Summary.loc[AW_Summary['SOM'] == 'Shubham Kaushik'])["Geotagged Outlets"].sum()


geo_list = [only_abhijit_bera_column_g,only_abhimanyu_singh_column_g,only_abinash_patnaik_column_g,
only_amit_kumar_column_g,only_atin_kapoor_column_g,only_b_roy_column_g,
only_a_saxena_column_g,only_gokul_s_column_g,only_philip_p_column_g,
only_v_tandon_column_g,only_kasi_p_column_g,only_b_saha_column_g,
only_m_gurumurthy_column_g,only_manoj_razdan_column_g,only_m_bhatt_column_g,
only_m_rahman_column_g,only_n_pandya_column_g,
only_n_kumar_column_g,only_n_sood_column_g,only_p_pandey_column_g,
only_r_samatarey_column_g,only_r_kumar_A_column_g,only_r_mathew_column_g,
only_s_r_column_g,only_sanjeev_k_column_g,only_sarvanan_column_g,
only_s_mitra_column_g,only_s_nagpal_column_g,only_thanis_d_column_g,
only_m_kumar_column_g,only_v_sharma_column_g,only_lalit_joshi_column_g,
only_s_chopra_column_g,only_s_sharma_column_g, only_s_kaushik_column_g]


# BASE OUTLETS

only_abhijit_bera_column_ = (AW_Summary.loc[AW_Summary['SOM'] == "Abhijit Bera"])["Base Outlets"].sum()
only_abhimanyu_singh_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Abhimanyu Singh'])["Base Outlets"].sum()
only_abinash_patnaik_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Abinash Patnaik'])["Base Outlets"].sum()
only_amit_kumar_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Amit Kumar Dhage'])["Base Outlets"].sum() 
only_atin_kapoor_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Atin Kapoor'])["Base Outlets"].sum()

only_b_roy_column_= (AW_Summary.loc[AW_Summary['SOM'] == 'Biswapriyo Roy'])["Base Outlets"].sum()

only_a_saxena_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'ABHISHEK SAXENA'])["Base Outlets"].sum()

only_gokul_s_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'GOKUL S'])["Base Outlets"].sum()

only_philip_p_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Philip P Koshy'])["Base Outlets"].sum()

only_v_tandon_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Vikas Tandon'])["Base Outlets"].sum()

only_kasi_p_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'KASI VISHWANATHAN P'])["Base Outlets"].sum()

only_b_saha_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Biswarup Saha'])["Base Outlets"].sum()

only_m_gurumurthy_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Mahesh Gurumurthy'])["Base Outlets"].sum()

only_manoj_razdan_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Manoj Razdan ' ])["Base Outlets"].sum()

only_m_bhatt_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Minesh Bhatt'])["Base Outlets"].sum()

only_m_rahman_column_= (AW_Summary.loc[AW_Summary['SOM'] == 'Musir Rahman'])["Base Outlets"].sum()

only_n_pandya_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Nihar Pandya'])["Base Outlets"].sum()

only_n_kumar_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Nilesh Kumar'])["Base Outlets"].sum()

only_n_sood_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Nitesh Sood'])["Base Outlets"].sum()

only_p_pandey_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Pardeep Pandey'])["Base Outlets"].sum()

only_r_samatarey_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Rajesh  Samantaray'])["Base Outlets"].sum()

only_r_kumar_A_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'RAMESH KUMAR A'])["Base Outlets"].sum()

only_r_mathew_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Rejo Mathew'])["Base Outlets"].sum()

only_s_r_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'SAIKANNAN R'])["Base Outlets"].sum()

only_sanjeev_k_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Sanjeev Kumar'])["Base Outlets"].sum()

only_sarvanan_column_ =(AW_Summary.loc[AW_Summary['SOM'] == 'Saravanan TN'])["Base Outlets"].sum()

only_s_mitra_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Sudipta Mitra'])["Base Outlets"].sum()

only_s_nagpal_column_ =(AW_Summary.loc[AW_Summary['SOM'] == 'Sumeet Nagpal'])["Base Outlets"].sum()

only_thanis_d_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'THANIS D'])["Base Outlets"].sum()

only_m_kumar_column_ =(AW_Summary.loc[AW_Summary['SOM'] == 'Mihir Kumar'])["Base Outlets"].sum()

only_v_sharma_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Vishal Sharma'])["Base Outlets"].sum()

only_lalit_joshi_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Lalit Mohan Joshi'])["Base Outlets"].sum()

only_s_chopra_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Sanjeev Chopra'])["Base Outlets"].sum()

only_s_sharma_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Shailendra Sharma'])["Base Outlets"].sum()
only_s_kaushik_column_ = (AW_Summary.loc[AW_Summary['SOM'] == 'Shubham Kaushik'])["Base Outlets"].sum()

base_list = [only_abhijit_bera_column_,only_abhimanyu_singh_column_,only_abinash_patnaik_column_,
only_amit_kumar_column_,only_atin_kapoor_column_,only_b_roy_column_,
only_a_saxena_column_,only_gokul_s_column_,only_philip_p_column_,
only_v_tandon_column_,only_kasi_p_column_,only_b_saha_column_,
only_m_gurumurthy_column_,only_manoj_razdan_column_,only_m_bhatt_column_,
only_m_rahman_column_,only_n_pandya_column_,
only_n_kumar_column_,only_n_sood_column_,only_p_pandey_column_,
only_r_samatarey_column_,only_r_kumar_A_column_,only_r_mathew_column_,
only_s_r_column_,only_sanjeev_k_column_,only_sarvanan_column_,
only_s_mitra_column_,only_s_nagpal_column_,only_thanis_d_column_,
only_m_kumar_column_,only_v_sharma_column_,only_lalit_joshi_column_,
only_s_chopra_column_,only_s_sharma_column_, only_s_kaushik_column_]



#for AW code
q= AW_Summary['SOM'].value_counts()['Abhijit Bera']
w = AW_Summary['SOM'].value_counts()['Abhimanyu Singh']
e = AW_Summary['SOM'].value_counts()['Abinash Patnaik']
r = AW_Summary['SOM'].value_counts()['Amit Kumar Dhage']
t = AW_Summary['SOM'].value_counts()['Atin Kapoor']
y = AW_Summary['SOM'].value_counts()['Biswapriyo Roy']
u = AW_Summary['SOM'].value_counts()['ABHISHEK SAXENA']
i = AW_Summary['SOM'].value_counts()['GOKUL S']
o = AW_Summary['SOM'].value_counts()['Philip P Koshy']
p = AW_Summary['SOM'].value_counts()['Vikas Tandon']
a = AW_Summary['SOM'].value_counts()['KASI VISHWANATHAN P']
s = AW_Summary['SOM'].value_counts()['Biswarup Saha']
d = AW_Summary['SOM'].value_counts()['Mahesh Gurumurthy']
f = AW_Summary['SOM'].value_counts()['Manoj Razdan ']
g = AW_Summary['SOM'].value_counts()['Minesh Bhatt']
h = AW_Summary['SOM'].value_counts()['Musir Rahman']
j = AW_Summary['SOM'].value_counts()['Nihar Pandya']
k = AW_Summary['SOM'].value_counts()['Nilesh Kumar']
l = AW_Summary['SOM'].value_counts()['Nitesh Sood']
z = AW_Summary['SOM'].value_counts()['Pardeep Pandey']
x = AW_Summary['SOM'].value_counts()['Rajesh  Samantaray']
c = AW_Summary['SOM'].value_counts()['RAMESH KUMAR A']
v = AW_Summary['SOM'].value_counts()['Rejo Mathew']
b = AW_Summary['SOM'].value_counts()['SAIKANNAN R']
n = AW_Summary['SOM'].value_counts()['Sanjeev Kumar']
m = AW_Summary['SOM'].value_counts()['Saravanan TN']
qq = AW_Summary['SOM'].value_counts()['Sudipta Mitra']
ww = AW_Summary['SOM'].value_counts()['Sumeet Nagpal']
ee = AW_Summary['SOM'].value_counts()['THANIS D']
rr = AW_Summary['SOM'].value_counts()['Mihir Kumar']
tt = AW_Summary['SOM'].value_counts()['Vishal Sharma']
mm = AW_Summary['SOM'].value_counts()['Lalit Mohan Joshi']
yy = AW_Summary['SOM'].value_counts()['Sanjeev Chopra']
uu = AW_Summary['SOM'].value_counts()['Shailendra Sharma']
ii = AW_Summary['SOM'].value_counts()['Shubham Kaushik']
aw_list = [q,w,e,r,t,y,u,i,o,p,a,s,d,f,g,h,j,k,l,z,x,c,v,b,n,m,qq,ww,ee,rr,tt,mm,yy,uu,ii]



#for re-tag
only_abhijit_bera_column_r = retag_file_combine['SOM'].value_counts()['Abhijit Bera']
only_abhimanyu_singh_column_r = retag_file_combine['SOM'].value_counts()['Abhimanyu Singh']
only_abinash_patnaik_column_r = retag_file_combine['SOM'].value_counts()['Abinash Patnaik']
only_amit_kumar_column_r = retag_file_combine['SOM'].value_counts()['Amit Kumar Dhage'] 
only_atin_kapoor_column_r = retag_file_combine['SOM'].value_counts()['Atin Kapoor']

only_b_roy_column_r= retag_file_combine['SOM'].value_counts()['Biswapriyo Roy']

only_a_saxena_column_r = retag_file_combine['SOM'].value_counts()['ABHISHEK SAXENA']

only_gokul_s_column_r = retag_file_combine['SOM'].value_counts()['GOKUL S']

only_philip_p_column_r = retag_file_combine['SOM'].value_counts()['Philip P Koshy']

only_v_tandon_column_r = retag_file_combine['SOM'].value_counts()['Vikas Tandon']

only_kasi_p_column_r = retag_file_combine['SOM'].value_counts()['KASI VISHWANATHAN P']

only_b_saha_column_r = retag_file_combine['SOM'].value_counts()['Biswarup Saha']

only_m_gurumurthy_column_r = retag_file_combine['SOM'].value_counts()['Mahesh Gurumurthy']

only_manoj_razdan_column_r = retag_file_combine['SOM'].value_counts()['Manoj Razdan ']

only_m_bhatt_column_r = retag_file_combine['SOM'].value_counts()['Minesh Bhatt']

only_m_rahman_column_r = retag_file_combine['SOM'].value_counts()['Musir Rahman']

only_n_pandya_column_r = retag_file_combine['SOM'].value_counts()['Nihar Pandya']

only_n_kumar_column_r = retag_file_combine['SOM'].value_counts()['Nilesh Kumar']

only_n_sood_column_r = retag_file_combine['SOM'].value_counts()['Nitesh Sood']

only_p_pandey_column_r = retag_file_combine['SOM'].value_counts()['Pardeep Pandey']

only_r_samatarey_column_r = retag_file_combine['SOM'].value_counts()['Rajesh  Samantaray']

only_r_kumar_A_column_r = retag_file_combine['SOM'].value_counts()['RAMESH KUMAR A']

only_r_mathew_column_r = retag_file_combine['SOM'].value_counts()['Rejo Mathew']

only_s_r_column_r = retag_file_combine['SOM'].value_counts()['SAIKANNAN R']

only_sanjeev_k_column_r = retag_file_combine['SOM'].value_counts()['Sanjeev Kumar']

only_sarvanan_column_r = retag_file_combine['SOM'].value_counts()['Saravanan TN']

only_s_mitra_column_r = retag_file_combine['SOM'].value_counts()['Sudipta Mitra']

only_s_nagpal_column_r = retag_file_combine['SOM'].value_counts()['Sumeet Nagpal']

only_thanis_d_column_r = retag_file_combine['SOM'].value_counts()['THANIS D']

only_m_kumar_column_r = retag_file_combine['SOM'].value_counts()['Mihir Kumar']

only_v_sharma_column_r = retag_file_combine['SOM'].value_counts()['Vishal Sharma']

only_lalit_joshi_column_r = retag_file_combine['SOM'].value_counts()['Lalit Mohan Joshi']

only_s_chopra_column_r = retag_file_combine['SOM'].value_counts()['Sanjeev Chopra']

only_s_sharma_column_r = retag_file_combine['SOM'].value_counts()['Shailendra Sharma']
only_s_kaushik_column_r = retag_file_combine['SOM'].value_counts()['Shubham Kaushik']

retag_list = [only_abhijit_bera_column_r,only_abhimanyu_singh_column_r,only_abinash_patnaik_column_r,
only_amit_kumar_column_r,only_atin_kapoor_column_r,only_b_roy_column_r,
only_a_saxena_column_r,only_gokul_s_column_r,only_philip_p_column_r,
only_v_tandon_column_r,only_kasi_p_column_r,only_b_saha_column_r,
only_m_gurumurthy_column_r,only_manoj_razdan_column_r,only_m_bhatt_column_r,
only_m_rahman_column_r,only_n_pandya_column_r,
only_n_kumar_column_r,only_n_sood_column_r,only_p_pandey_column_r,
only_r_samatarey_column_r,only_r_kumar_A_column_r,only_r_mathew_column_r,
only_s_r_column_r,only_sanjeev_k_column_r,only_sarvanan_column_r,
only_s_mitra_column_r,only_s_nagpal_column_r,only_thanis_d_column_r,
only_m_kumar_column_r,only_v_sharma_column_r,only_lalit_joshi_column_r,
only_s_chopra_column_r,only_s_sharma_column_r, only_s_kaushik_column_r]

Som_dataframee["Geotag outlets"] = geo_list
Som_dataframee["Total outlets"] = base_list
Som_dataframee["AW count"] = aw_list
Som_dataframee["retag outlets"] = retag_list


# dashboard and Som based table merger

def multiple_dfs(df_list, sheets, file_name, spaces):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    row = 0
    for dataframe in df_list:
        dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)   
        row = row + len(dataframe.index) + spaces + 1
    writer.save()

# list of dataframes
dfs = [Dashboard,Som_dataframee]

# run function
multiple_dfs(dfs, 'Dashboard', "C:\\Users\\91859\\Fast_docs\\Dashboard_output.xlsx", 2)

main("C:\\Users\\91859\\Fast_docs\\Final Output") 
"""






