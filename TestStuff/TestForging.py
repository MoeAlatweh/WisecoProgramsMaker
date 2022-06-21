
import pandas as pd

forging_file_path = 'H:\CNC_Programming\Moe A\WisecoApplications\TestStuff\Forging_Data_EMSS_vs_Model.xlsx'

forging_file = pd.read_excel(forging_file_path, sheet_name=None)

# print(forging_file)
#
# print(forging_file["Sheet1"]['Forging'])

# MAKE for LOOP TO STORE ALL FORGING NUMBERS IN ONE LIST
forging_list = []
for forging in forging_file['Forging_Sheet']['Forging']:
    forging_list.append(forging)
print("Forging List:", forging_list)
print("Forging List Quantity:", len(forging_list))

forging_number = "F6048TDX"
print("Forging Number:", forging_number)
# MAKE for LOOP TO FIND index OF FORGING BY ENTER THE (forging_number) AND ITERATE IT INSIDE THE (forging_list)
# WE NEED THE index TO CAN ACCESS ANY DIMENSION WE WANT
if forging_number in forging_list:
    forging_number_index = forging_list.index(forging_number)
    print("Forging Number Index:", forging_number_index)
    # TO ACCESS AND GET "FORGING OD" BY USE SHEET NAME(FROM EXCEL FILE), FORGING index(FROM ABOVE for LOOP),
    # AND COLUMN NAME OF DIMENSION(FROM EXCEL FILE)
    print("Forge O.D:", forging_file["Forging_Sheet"].at[forging_number_index, '(B) Forge O.D.'])
    print("Boss Insd Spacing:", forging_file["Forging_Sheet"].at[forging_number_index, '(J) Boss Insd Spacing'])

else:
    print("FORGING DOES NOT EXIST")

# TO ACCESS AND GET "FORGING OD" BY USE SHEET NAME(FROM EXCEL FILE), FORGING index(FROM ABOVE for LOOP),
# AND COLUMN NAME OF DIMENSION(FROM EXCEL FILE)
# print(forging_file["Sheet1"].at[forging_number_index, '(B) Forge O.D.'])


# MIN_DOME(A)  FORGE_OD(B)    FORGE_REF_LENGTH(F)   BOSS_INSIDE_SPACING(J)
# RING_BELT_ID(S)     BOSS_OUTSIDE_SPACING(U)     OUTSIDE_RING_BELT_HT(X)
