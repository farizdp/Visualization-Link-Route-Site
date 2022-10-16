from diagrams import Diagram, Edge
from diagrams.custom import Custom
from itertools import combinations
import pandas as pd
import os, time
pd.options.mode.chained_assignment = None

path = os.getcwd()
doc = path  + '/doc/'
icon = path + '/icon/'
files = os.listdir(doc)
files_xls = [doc + f for f in files if f[-4:] == 'xlsx']
latest_file = max(files_xls, key=os.path.getctime)

tabs = pd.ExcelFile(latest_file).sheet_names 
if ('Master LR' in tabs):
    data = pd.read_excel(latest_file, "Master LR")
    data_tx = data.iloc[:, [5,9,10,11,12,13,17,19,67,68]]
elif ('LINKROUTE' in tabs):
    data = pd.read_excel(latest_file, "LINKROUTE")
    data_raw = data.iloc[:, [3,9,10,11,12,13,17,19,53]].drop_duplicates()
    data_tx = data_raw[(data_raw['LINK BW'] != '-') & (data_raw['LINK BW'] != 0)]
    temp = data_tx[['Usage Link Existing / 65%','LINK BW']].astype(float)
    data_tx['Util'] = temp['Usage Link Existing / 65%'].div(temp['LINK BW'])
data_tx = data_tx.dropna()

site_simpul = icon + 'tower_1.png'
end_site = icon + 'tower_2.png'
metroe = icon + 'metroe.png'

site = "INPUT SITE ID"

# Create diagrams with direct near end
ne = data_tx[data_tx["SITE ID NE"] == site]
df_ne = ne.groupby(['SITE ID FE','LINK OWNER']).last().reset_index()
fe = data_tx[data_tx["SITE ID FE"] == site]
df_fe = fe.groupby(['SITE ID NE','LINK OWNER']).last().reset_index()
df_fe.rename(columns = {'SITE ID NE': 'SITE ID FE', 'SITE ID FE': 'SITE ID NE'}, inplace = True)
frames = [df_ne, df_fe]
df_result = pd.concat(frames)

with Diagram(site + ' ' + df_ne['SITE NAME NE'].values[0], graph_attr = {"labelloc":"t"}, show = False, direction = "LR"):
    near_end = Custom(df_ne["SITE ID NE"][0], site_simpul)
    for index, row in df_result.iterrows():
        tower = metroe if row["SITE ID FE"] == 'Metro-E' else end_site
        far_end = Custom(row['SITE ID FE'], tower)
        a = str(row['LINK BW']) + ' / ' + str("{:.2f}".format(row['Util']*100)) + '%'
        b = "Blue" if row['LINK OWNER'] == 'TELKOM' else "Red"
        near_end << Edge(label = a, color = b, style = "bold") >> far_end

# Create diagram with 1 full link route
filtered = data_tx[data_tx["SITE ID"] == site]
sitelist = filtered['SITE ID NE'].tolist() + filtered['SITE ID FE'].tolist()
mylist = list(dict.fromkeys(sitelist))
comb = list(combinations(mylist, 2))

with Diagram(site + ' ' + df_ne['SITE NAME NE'].values[0], graph_attr = {"labelloc":"t"}, show = False, direction = "LR"):
    twr = []
    for a in mylist:
        tower = metroe if a == 'Metro-E' else end_site
        temp = Custom(a, tower)
        twr.append(temp)
    for b in comb:
        link = filtered[(filtered['SITE ID NE'] == b[0]) & (filtered['SITE ID FE'] == b[1])]
        if link.empty == False:
            x = str(link['LINK BW'].values[0]) + ' / ' + str("{:.2f}".format(link['Util'].values[0]*100)) + '%'
            y = "Blue" if link['LINK OWNER'].values[0] == 'TELKOM' else "Red"
            twr[mylist.index(b[0])] << Edge(label = x, color = y, style = "bold") >> twr[mylist.index(b[1])]