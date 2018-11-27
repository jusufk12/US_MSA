import pandas as pd
import os
path = "C:\\Users\\Datasoft\\Desktop\\MSA_files"

files = os.listdir(path)

mse = pd.read_excel("C:\\Users\\Datasoft\\Desktop\\New folder (2)\\MSAs.xlsx")
mse = pd.DataFrame(mse['MSA'].str.split(',', 1).tolist(), columns=['flips', 'row'])
#filtered_msa_names = mse['MSA'].tolist()
filtered_msa_names = mse['flips'].tolist()
final = pd.DataFrame()
writer = pd.ExcelWriter(path+'\\MSA_final.xlsx')
for i in range(len(files)):
    new_columns = []
    new_columns.append("Geo_NAME")
    df = pd.read_csv(path+"\\"+files[i], encoding="ISO-8859-1")
    df.iloc[len(df)-1] = df.columns
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df = df.reset_index(drop=True)
    for n in range(len(df.columns)):
        if df.columns[n][-1:] != "s":
            if "Geo" not in df.columns[n]:
                new_columns.append(df.columns[n])
    df = df[new_columns]

    df['Geo_NAME'] = df['Geo_NAME'].str.replace(" Metro Area", "")
    df['Geo_NAME'] = df['Geo_NAME'].str.replace(" Micro Area", "")
    dataframe = pd.DataFrame(columns=df.columns)
    for m in range(len(filtered_msa_names)):
        print(filtered_msa_names[m]+", "+str(mse.loc[mse.index == m].values[0][1]))
        filtered_msa_not_found = []
        aa = 0
        for g in range(len(df)):
            if filtered_msa_names[m] == df['Geo_NAME'][g].split(",")[0]:
                if mse.loc[mse.index == m].values[0][1] == df['Geo_NAME'][g].split(",")[1]:
                    a = 1

                    print("\t"+df['Geo_NAME'][g])
                    new_row = df.loc[df.index == g]
                    dataframe = dataframe.append(new_row, ignore_index=True)
                else:
                    print("\t\t\t"+df['Geo_NAME'][g])

    year = dataframe.columns[int((len(dataframe)-1)/2)].split("_")[0]
    year2 = "20"+year[3:]
    if year2 == '2009':
       a = df.loc[df.index == 700]
       dataframe = dataframe.append(a, ignore_index=True)

    if year2 == '2012':
        a = df.loc[df.index==61]
        dataframe = dataframe.append(a, ignore_index=True)

    if year2 == '2011':
        a = df.loc[df.index==61]
        dataframe = dataframe.append(a, ignore_index=True)
    if year2 == '2010':
        a = df.loc[df.index==61]
        dataframe = dataframe.append(a, ignore_index=True)

    last_row = df.loc[df.index ==len(df)-1]
    dataframe = dataframe.append(last_row, ignore_index=True)

    idustry = []
    eductation = []
    for j in range(len(dataframe.columns)):
        if "C24050" in dataframe.columns[j]:
            idustry.append(dataframe.columns[j])
        elif "T025" in dataframe.columns[j]:
            eductation.append(dataframe.columns[j])
    columns = ['Geo_NAME',year+'_5yr_B01003001', year+'_5yr_B25001001', year+'_5yr_B01001012', year+'_5yr_B01001013', year+'_5yr_B01001014', year+'_5yr_B01001036', year+'_5yr_B01001037', year+'_5yr_B01001038', 'SE_T033_006']
    final_cols = columns + idustry + eductation
    dataframe = dataframe[final_cols]
    dataframe.columns = dataframe.iloc[len(dataframe)-1]
    dataframe = dataframe.drop(dataframe.index[len(dataframe)-1])




    print(year2)
    dataframe.to_excel(writer, year2, index=False)
writer.save()
writer.close()





