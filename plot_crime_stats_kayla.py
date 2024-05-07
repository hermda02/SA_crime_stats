import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os
import numpy as np

def parse_q1(file):
    df = pd.read_excel(file, sheet_name="Tables 2020 Quarterly")

    labels = df.iloc[21:26, 2]
    dates = ['2018-2019 Q1', '2019-2020 Q1', '2020-2021 Q1', '2021-2022 Q1', '2022-2023 Q1',]#df.iloc[10, 3:8]
    data = df.iloc[21:26, 3:8]

    data = relabel_data(data, dates, labels)

    return data

def parse_q2(file):
    df = pd.read_excel(file, sheet_name="Tables 2022 Quarter 2")

    labels = df.iloc[21:26, 2]
    dates = ['2018-2019 Q2', '2019-2020 Q2', '2020-2021 Q2', '2021-2022 Q2', '2022-2023 Q2',]#df.iloc[10, 3:8]
    data = df.iloc[21:26, 8:13]

    data = relabel_data(data, dates, labels)

    return data
    
def parse_q3(file):
    df = pd.read_excel(file, sheet_name="Crime stats per component")

    data = df.iloc[22:26, 3:8]

    # sums = data.sum()

    data.loc["26"] = data.sum()
    
    # append sums to the data frame
    labels = df.iloc[22:26, 2]
    dates = ['2018-2019 Q3', '2019-2020 Q3', '2020-2021 Q3', '2021-2022 Q3', '2022-2023 Q3',]#df.iloc[10, 3:8]
    labels.loc['26'] = "Total Sexual Offences"
    
    data = relabel_data(data, dates, labels)
    
    return data
    
def parse_q4(file):
    df = pd.read_excel(file, sheet_name="Crime stats per component")

    data = df.iloc[22:26, 3:8]
    
    data.loc["26"] = data.sum()
    
    # append sums to the data frame
    labels = df.iloc[22:26, 2]
    dates = ['2018-2019 Q4', '2019-2020 Q4', '2020-2021 Q4', '2021-2022 Q4', '2022-2023 Q4',]#df.iloc[10, 3:8]
    labels.loc['26'] = "Total Sexual Offences"
    
    data = relabel_data(data, dates, labels)

    return data


def relabel_data(data, dates, labels):
    new_labels_y = {}
    new_labels_x = {}

    for i in range(len(dates)):
        new_labels_x[data.columns.values[i]] = dates[i].lower().title()

    data = data.rename(columns=new_labels_x)

    data = data.T

    for i in range(len(labels.values)):
        new_labels_y[data.columns.values[i]] = labels.values[i].lower().title()

    data = data.rename(columns=new_labels_y)

    return data

def swap_and_sort(df):
    inds = df.index()
    
def quarter_to_date(quarter: str):
    if 'Q1' in quarter:
        out = 'April-June ' + quarter[:4] 
    if 'Q2' in quarter:
        out = 'July-September ' + quarter[:4] 
    if 'Q3' in quarter:
        out = 'October-December ' + quarter[:4] 
    if 'Q4' in quarter:
        out = 'January-March ' + quarter[5:9] 

    return out
        
if __name__ == "__main__":
    files = ['2022-2023-Q1-crime-stats.xlsx', '2022-2023-Q2-crime-stats.xlsx','2022-2023-Q3-crime-stats.xlsx','2022-2023-Q4-crime-stats.xlsx',]
    q1 = parse_q1(files[0])
    q2 = parse_q2(files[1])
    q3 = parse_q3(files[2])
    q4 = parse_q4(files[3])
    cat_dat = pd.concat([q1, q2, q3, q4]).sort_index()

    inds = cat_dat.index

    quarter_rename = {}
    for ind in inds:
        quarter_rename[ind] = quarter_to_date(ind)

    cat_dat = cat_dat.rename(index=quarter_rename)
    
    lockdown_data = {
        "March 27 - April 30 2020": 5,
        "May 1 - May 31 2020": 4,
        "June 1 - August 17 2020": 3,
        "August 18 - September 20 2020": 2,
        "September 21 - December 28 2020": 1,
        "Decemeber 29 - February 28 2021": 3,
        "March 1 - March 31 2021": 1,
    }

    colors = ['red', 'orange', 'yellow', 'yellowgreen', 'green']
    
    handles = [mpatches.Patch(color=c) for c in colors]
    levels = [5,4,3,2,1]
    
    ax = cat_dat.plot(style='--', figsize=(10,5))
    leg1 = ax.legend(handles, levels, loc='upper left', title='Lockdown Level')
    ax.set_xticks(range(len(cat_dat)))
    ax.set_xticklabels(cat_dat.index.tolist(), rotation=90)
    ax.axvspan(7.85, 8.33, alpha=0.5,color=colors[0])
    ax.axvspan(8.33, 8.66, alpha=0.5,color=colors[1])
    ax.axvspan(8.66, 9.5, alpha=0.5, color=colors[2])
    ax.axvspan(9.5, 9.9, alpha=0.5, color=colors[3])
    ax.axvspan(9.9, 11, alpha=0.5, color=colors[4])
    ax.axvspan(11, 11.66, alpha=0.5, color=colors[2])
    ax.axvspan(11.66, 12, alpha=0.5, color=colors[4])
    # ax.text(8,3000, 'Lockdown Level: 5', color='k', rotation=90, fontsize=12)
    # ax.text(7.33,3000, 'Lockdown Level: 5', color='k', rotation=90, fontsize=12)
    ax.set_ylim([0,16000])
    plt.legend(loc='center left')
    ax.add_artist(leg1)
    plt.savefig('SA_crime_stats', dpi=100, bbox_inches='tight')
    # plt.show()
