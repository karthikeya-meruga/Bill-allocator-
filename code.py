#%%
import pandas
import os

excel_name = 'NWA_Mapping_Sample.xlsx'
nwa_name ='NWAMasterCode'
emp_name = 'Mapping'

nwa = pandas.read_excel(excel_name,sheet_name=nwa_name)
emp = pandas.read_excel(excel_name,sheet_name=emp_name)

emp['NWA Code']=''

nwa = nwa.rename(columns= lambda x: x.strip())
emp = emp.rename(columns= lambda x: x.strip())
emp = emp.sort_values(by='Total cost', ascending=False)
nwa = nwa.sort_values(by='Available', ascending=False)
# %%
nwa
# %%
emp
# %%
nwa_marker=0
nwa_list=list(nwa['NWA Code'])
nwa_item=nwa_list[nwa_marker]
nwa_item
# %%
def allocate(balance, emp_no):
    global nwa_marker
    global nwa_list
    nwa_item = nwa_list[nwa_marker]
    nwa_balance = float(nwa.loc[nwa['NWA Code']==nwa_item, 'Available'].values[0])
    print(f"    NWA Code: {nwa_item}    NWA Balance: {nwa_balance}")

    if balance <= nwa_balance:
        # change and update nwa_balance
        nwa_balance = nwa_balance-balance
        nwa.loc[nwa['NWA Code']== nwa_item, 'Available'] = nwa_balance

        update=str(emp.loc[emp['Emp No']==emp_no,'NWA Code'].values[0])
        if update != '':
            update = update+'\n'
        update=update+", "+str(nwa_item)+':'+str(balance)
        emp.loc[emp['Emp No']==emp_no,'NWA Code'] = update
        balance=0
        return balance

    else:
        balance = balance-nwa_balance
        update=str(emp.loc[emp['Emp No']==emp_no,'NWA Code'].values[0])
        if update != '':
            update = update+'\n'
        update=update+", "+str(nwa_item)+':'+str(nwa_balance)
        emp.loc[emp['Emp No']==emp_no,'NWA Code'] = update

        nwa_balance=0
        nwa_marker = nwa_marker + 1
        nwa.loc[nwa['NWA Code']== nwa_item, 'Available'] = nwa_balance

        return balance
# %%
for emp_no in emp['Emp No']:
    balance = emp.loc[emp['Emp No']==emp_no, 'Total cost'].values[0]
    while(balance>0):
        print(f"Employee Number= {emp_no}, balance={balance}")
        balance = allocate(balance,emp_no)
        emp.loc[emp['Emp No']==str(emp_no),'Total cost']=balance

# %%
emp
# %%
nwa
# %%
emp.to_excel(r'./output.xlsx', index=False, header=True)
# %%
