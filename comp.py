# a utility to compare key value pairs in different excel files, assuming that they have the same index codes in each.

import pandas as pd

df_a = pd.read_excel('FAMIS-JEs Record Center 3RD Qtr 2020.xlsx', sheet_name = None)
df_a = pd.read_excel('test.xlsx', sheet_name = None)
df_famis = pd.concat(df_a, ignore_index=True, sort=True)


data_top = df_famis.head()
data_top = list(data_top)

df_famis = df_famis.applymap(lambda s:s.lower().strip() if type(s) == str else s)

for top in data_top:

	for items in df_famis[top].dropna():
		if isinstance(items, str):
			if(items.lower() == 'index code'):
				df_famis.rename(columns={str(top):'index code'}, inplace = True)


				#print(top)
			elif(items.lower() == 'transaction amount'):
				df_famis.rename(columns={str(top):'transaction amount'}, inplace = True)



i = df_famis[df_famis['index code'] == 'index code'].index


j = df_famis[df_famis['transaction amount'] == 'transaction amount'].index



for headers in i:
	df_famis.at[ headers, 'index code' ] = None
	df_famis.at[ headers, 'transaction amount' ] = None

nec = ['index code', 'transaction amount']	

df_san = df_famis[nec].dropna()


san_tot = df_famis.groupby('index code')['transaction amount']#.sum()



san_tot = df_famis.groupby('index code')['transaction amount'].sum().round(2)


FAMIS_SET = san_tot.to_dict()

#########################################################
df_fy = pd.read_excel('Copy of FY.xlsx')

df_fy = df_fy.applymap(lambda s:s.lower().strip() if type(s) == str else s)

data_headers = df_fy.head()
data_headers = list(data_headers)



sum_tot = df_fy.groupby('INDEX CODE')['TOTAL']#.sum().round(2)
test_dict = {}


sum_tot = df_fy.groupby('INDEX CODE')['TOTAL'].sum().round(2)
FY_SET = sum_tot.to_dict()



for codes in FAMIS_SET:
	if FY_SET.get(codes):
		if FY_SET[codes] != FAMIS_SET[codes]:
			print(codes, "is wrong")
			print("FY VALUE IS: ", FY_SET[codes], "FOR INDEX CODE", codes)
			print("FAMIS VALUE IS: ", FAMIS_SET[codes], "FOR INDEX CODE", codes)

