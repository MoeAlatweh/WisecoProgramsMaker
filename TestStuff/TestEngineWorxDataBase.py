

import pyodbc
# TO CONNECT TO THE DATA BASE
# (SQL Server) : TYPE OF DRIVER
# (us-men-app-sql1) : NAME OF SERVER
# (EngineWorx) : NAME OF DATA BASE
# (Trusted_Connection=Yes) : SAME AUTHENTICATION INFO YOU USE TO LOG IN YOUR COMPUTER (Windows Authentication),
# IF THEY ARE DIFFERENT OR USE (SQL Authentication):
#   NEEDS TO ADD ('Uid=WISECOMANF\\DomainUser;') AND ('Pwd='YourPasswordToSQL;') with ('Trusted_Connection=No;')
# TO BE ABLE TO ACCESS THE DATABASE, NEEDS PERMISSION AND DATABASE INFO FROM IT DEPARTMENT.

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=us-men-app-sql1;'
                      'Database=EngineWorx;'
                      'Trusted_Connection=Yes;')

# CREATE CURSOR TO POINT THE DATA
cursor = conn.cursor()

# TO PRINT LOGIN INFO FOR THE DATABASE
cursor.execute("SELECT * FROM master.sys.sql_logins")
for i in cursor:
    print(i)
print()
# TO PRINT MORE LOGIN INFO FOR THE DATABASE
cursor.execute("SELECT * FROM master.sys.syslogins")
for i in cursor:
    print(i)
print()

# TO PRINT USER'S RELATED INFO FOR THE DATABASE (IF EVER NEEDED)
# cursor.execute("SELECT * FROM master.sys.database_principals")
# for i in cursor:
#     print(i)
# print()
# TO PRINT USER'S RELATED INFO FOR THE DATABASE (IF EVER NEEDED)
# cursor.execute("SELECT * FROM master.sys.sysusers")
# for i in cursor:
#     print(i)
# print()

# region  <<<<============================[Tables Names]============================>>>>
# TO GET TABLES NAMES OF THE DATA BASE
cursor.execute("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'")
tables_names = []
for i in cursor:
    # print(i)
    tables_names.append(i)
print("[#]Tables Names of EngineWorx DataBase :")
print(*tables_names, sep="\n")
print()
# endregion  <<<<===========================[Tables Names]===========================>>>>

# region  <<<<========================[Columns Names of <SpexPiston> Table]========================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<======================[Columns Names of <SpexPiston> Table]======================>>>>

# region  <<<<====================[Columns Names of <SpexPiston_PinBore> Table]====================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_PinBore'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_PinBore> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==================[Columns Names of <SpexPiston_PinBore> Table]===================>>>>

# region  <<<<==================[Columns Names of <SpexPiston_OilDrainHole> Table]==================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_OilDrainHole'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_OilDrainHole> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<================[Columns Names of <SpexPiston_OilDrainHole> Table]=================>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpOilDrainHoleTypes> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpOilDrainHoleTypes'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpOilDrainHoleTypes> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpOilDrainHoleTypes> Table]==============>>>>

# region  <<<<==================[Columns Names of <SpexPiston_SemiFinishTurn> Table]==================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_SemiFinishTurn'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_SemiFinishTurn> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<================[Columns Names of <SpexPiston_SemiFinishTurn> Table]=================>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpLatheDomeTypes> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpLatheDomeTypes'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpLatheDomeTypes> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpLatheDomeTypes> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpPressureSeal> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpPressureSeal'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpPressureSeal> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpPressureSeal> Table]==============>>>>

# region  <<<<==================[Columns Names of <SpexPiston_Milling> Table]==================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_Milling'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_Milling> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<================[Columns Names of <SpexPiston_Milling> Table]=================>>>>

# region  <<<<================[Columns Names of <SpexPiston_IDBore> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_IDBore'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_IDBore> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_IDBore> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_Stamping> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_Stamping'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_Stamping> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_Stamping> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_FinishTurn> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_FinishTurn'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_FinishTurn> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_FinishTurn> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpEdgeBreakSkirtTop> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpEdgeBreakSkirtTop'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpEdgeBreakSkirtTop> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpEdgeBreakSkirtTop> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpEdgeBreakSkirtBottom> Table]================>>>>
# To get column names from table
cursor.execute(
    "SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpEdgeBreakSkirtBottom'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpEdgeBreakSkirtBottom> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpEdgeBreakSkirtBottom> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpGasPorts> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpGasPorts'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpGasPorts> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpGasPorts> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_GasPortsPistonPinOiling> Table]================>>>>
# To get column names from table
cursor.execute(
    "SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_GasPortsPistonPinOiling'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_GasPortsPistonPinOiling> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_GasPortsPistonPinOiling> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpPressureFedOilHoleType> Table]================>>>>
# To get column names from table
cursor.execute(
    "SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpPressureFedOilHoleType'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpPressureFedOilHoleType> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpPressureFedOilHoleType> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_LkpPressureFedOilHoleSlots> Table]================>>>>
# To get column names from table
cursor.execute(
    "SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_LkpPressureFedOilHoleSlots'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_LkpPressureFedOilHoleSlots> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_LkpPressureFedOilHoleSlots> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_PistonPinHoleFinish> "Honing" Table]================>>>>
# To get column names from table
cursor.execute(
    "SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_PistonPinHoleFinish'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_PistonPinHoleFinish> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_PistonPinHoleFinish> "Honing" Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_Coating> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_Coating'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_Coating> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_Coating> Table]==============>>>>


# region  <<<<================[Columns Names of <SpexWristPins> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexWristPins'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexWristPins> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexWristPins> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexRetainerClips> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexRetainerClips'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexRetainerClips> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexRetainerClips> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexOilRailSupport> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexOilRailSupport'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexOilRailSupport> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexOilRailSupport> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexPiston_2Cycle> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexPiston_2Cycle'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexPiston_2Cycle> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexPiston_2Cycle> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexForge> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexForge'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexForge> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexForge> Table]==============>>>>

# region  <<<<================[Columns Names of <SpexForge_History> Table]================>>>>
# To get column names from table
cursor.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpexForge_History'")
column_names = []
for i in cursor:
    # print(i)
    column_names.append(i)
print("[#]Column Names of <SpexForge_History> Table :")
print(*column_names, sep="\n")
print()
# endregion  <<<<==============[Columns Names of <SpexForge_History> Table]==============>>>>


# region  <<<<============================[Spex Information from DataBase]============================>>>>
# <<<<<<<<------------------------------------------------------------------------------------------->>>>>>>>>>
#      <<<<<<<<--------------------------------------------------------------------------------->>>>>>>>>>
#          <<<<<<<<------------------------------------------------------------------------>>>>>>>>>>

print("<<<<============================[Spex Information from DataBase]============================>>>>")
# region  <<<<============================[Piston Job Numbers]============================>>>>
# Table Name : SpexPiston
# Column Name : Piston
cursor.execute('SELECT Piston FROM SpexPiston')
job_numbers = []
for i in cursor:
    # print(i)
    job_numbers.append(i)
print("[#]Job Numbers:")
# [25000:25003] : Range of List needed by index
print("     ", job_numbers[25000:25003])
# print(len(job_numbers))
# endregion  <<<<===========================[Piston Job Numbers]===========================>>>>

# region  <<<<============================[Piston Job ID Numbers]============================>>>>
# Table Name : SpexPiston
# Column Name : PistonID
cursor.execute('SELECT PistonID FROM SpexPiston')
job_id_numbers = []
for i in cursor:
    # print(i)
    job_id_numbers.append(i)
print("[#]Job ID Numbers:")
# [25000:25003] : Range of List needed by index
print("     ", job_id_numbers[25000:25003])
print("     ", job_id_numbers[-1])

print(len(job_id_numbers))
# endregion  <<<<===========================[Piston Job ID Numbers]===========================>>>>

# region  <<<<============================[EngineStrokeType]============================>>>>
# Table Name : SpexPiston
# Column Name : EngineStrokeType
cursor.execute('SELECT EngineStrokeType FROM SpexPiston')
engine_stroke_type = []
for i in cursor:
    # print(i)
    engine_stroke_type.append(i)
print("[#]EngineStrokeType:")
# [25000:25003] : Range of List needed by index
print("     ", engine_stroke_type[25000:25003])
# endregion  <<<<===========================[EngineStrokeType]===========================>>>>

# region  <<<<============================[Job Description]============================>>>>
# Table Name : SpexPiston
# Column Name : Description
cursor.execute('SELECT Description FROM SpexPiston')
job_description = []
for i in cursor:
    # print(i)
    job_description.append(i)
print("[#]Job Description:")
# [25000:25003] : Range of List needed by index
print("     ", job_description[25000:25003])
# endregion  <<<<===========================[Description]===========================>>>>

# region  <<<<============================[Released Status]============================>>>>
# Table Name : SpexPiston
# Column Name : Released_Y_N
cursor.execute('SELECT Released_Y_N FROM SpexPiston')
released_status = []
for i in cursor:
    # print(i)
    released_status.append(i)
print("[#]Released Status:")
# [25000:25003] : Range of List needed by index
print("     ", released_status[25000:25003])
# endregion  <<<<===========================[Released Status]===========================>>>>

# region  <<<<============================[Date Released]============================>>>>
# Table Name : SpexPiston
# Column Name : DateReleased
cursor.execute('SELECT DateReleased FROM SpexPiston')
date_released = []
for i in cursor:
    # print(i)
    date_released.append(i)
print("[#]Date Released:")
# [25000:25003] : Range of List needed by index
print("     ", date_released[25000:25003])
# endregion  <<<<===========================[DateReleased]===========================>>>>

# region  <<<<============================[Released By]============================>>>>
# Table Name : SpexPiston
# Column Name : ReleasedBy
cursor.execute('SELECT ReleasedBy FROM SpexPiston')
released_by = []
for i in cursor:
    # print(i)
    released_by.append(i)
print("[#]Released By:")
# [25000:25003] : Range of List needed by index
print("     ", released_by[25000:25003])
# endregion  <<<<===========================[Released By]===========================>>>>

# region  <<<<============================[Wrist Pin Item ID -->> Not For Spec]============================>>>>
# Table Name : SpexWristPins
# Column Name : WristPinItemID
cursor.execute('SELECT WristPinItemID FROM SpexWristPins')
wrist_pin_item_id = []
for i in cursor:
    # print(i)
    wrist_pin_item_id.append(i)
print("[#]Wrist Pin Item ID:")
# [25000:25003] : Range of List needed by index
print("     ", wrist_pin_item_id[25000:25003])
# endregion  <<<<===========================[Wrist Pin Item ID -->> Not For Spex]===========================>>>>

# region  <<<<============================[WristPinID]============================>>>>
# Table Name : SpexPiston
# Column Name : WristPinID
cursor.execute('SELECT WristPinID FROM SpexPiston')
wrist_pin_id = []
for i in cursor:
    # print(i)
    wrist_pin_id.append(i)
print("[#]Wrist Pin ID:")
# [25000:25003] : Range of List needed by index
print("     ", wrist_pin_id[25000:25003])
# endregion  <<<<===========================[WristPinID]===========================>>>>

# region  <<<<============================[RetClip Item ID -->> Not For Spec]============================>>>>
# Table Name : SpexRetainerClips
# Column Name : RetClipItemID
cursor.execute('SELECT RetClipItemID FROM SpexRetainerClips')
ret_clip_item_id = []
for i in cursor:
    # print(i)
    ret_clip_item_id.append(i)
print("[#]RetClip Item ID:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_item_id[25000:25003])
# endregion  <<<<===========================[RetClip Item ID -->> Not For Spec]===========================>>>>

# region  <<<<============================[RetainerClipID]============================>>>>
# Table Name : SpexPiston
# Column Name : RetainerClipID
cursor.execute('SELECT RetainerClipID FROM SpexPiston')
retainer_clip_id = []
for i in cursor:
    # print(i)
    retainer_clip_id.append(i)
print("[#]Retainer Clip ID:")
# [25000:25003] : Range of List needed by index
print("     ", retainer_clip_id[25000:25003])
# endregion  <<<<===========================[RetainerClipID]===========================>>>>

# region  <<<<============================[Ring Item ID -->> XXXXX]============================>>>>
# Column Name : RingItemID
# cursor.execute('SELECT RingItemID FROM SpexPiston')
# ring_item_id = []
# for i in cursor:
#     # print(i)
#     ring_item_id.append(i)
# print("[#]RingItemID:")
# # [25000:25003] : Range of List needed by index
# print("     ", ring_item_id[25000:25003])
# endregion  <<<<===========================[Ring Item ID -->> XXXXX]===========================>>>>

# region  <<<<==========================[Oil Rail Support Item ID -->> Not For Spec]============================>>>>
# Table Name : SpexOilRailSupport
# Column Name : OilRailSupportItemID
cursor.execute('SELECT OilRailSupportItemID FROM SpexOilRailSupport')
oil_rail_support_item_id = []
for i in cursor:
    # print(i)
    oil_rail_support_item_id.append(i)
print("[#]Oil Rail Support Item ID:")
# [25000:25003] : Range of List needed by index
print("     ", oil_rail_support_item_id[25000:25003])
# endregion  <<<<=========================[Oil Rail Support Item ID -->> Not For Spec]==========================>>>>

# region  <<<<==========================[OilRailSupportID]============================>>>>
# Table Name : SpexPiston
# Column Name : OilRailSupportID
cursor.execute('SELECT OilRailSupportID FROM SpexPiston')
oil_rail_support_id = []
for i in cursor:
    # print(i)
    oil_rail_support_id.append(i)
print("[#]Oil Rail Support ID:")
# [25000:25003] : Range of List needed by index
print("     ", oil_rail_support_id[25000:25003])
# endregion  <<<<=========================[OilRailSupportID]==========================>>>>

# region  <<<<============================[ForgeSpecID]============================>>>>
# Table Name : SpexPiston
# Column Name : ForgeSpecID
cursor.execute('SELECT ForgeSpecID FROM SpexPiston')
forge_Spec_id = []
for i in cursor:
    # print(i)
    forge_Spec_id.append(i)
print("[#]Forge Spec ID:")
# [25000:25003] : Range of List needed by index
print("     ", forge_Spec_id[25000:25003])
# endregion  <<<<===========================[ForgeSpecID]===========================>>>>

# region  <<<<============================[Forge Item ID (Forging Number)]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeItemID
cursor.execute('SELECT ForgeItemID FROM SpexForge')
forge_item_id = []
for i in cursor:
    # print(i)
    forge_item_id.append(i)
print("[#]Forge Item ID (Forging Number):")
# [25000:25003] : Range of List needed by index
print("     ", forge_item_id[25000:25003])
# endregion  <<<<===========================[Forge Item ID (Forging Number)]===========================>>>>

# region  <<<<============================[Bore Diameter by Inch]============================>>>>
# Table Name : SpexPiston
# Column Name : Bore_IN
cursor.execute('SELECT Bore_IN FROM SpexPiston')
bore_diameter_by_inch = []
for i in cursor:
    # print(i)
    bore_diameter_by_inch.append(i)
print("[#]Bore Diameter by Inch:")
# [25000:25003] : Range of List needed by index
print("     ", bore_diameter_by_inch[25000:25003])
# endregion  <<<<===========================[Bore Diameter by Inch]===========================>>>>

# region  <<<<============================[Compression Height By Inch]============================>>>>
# Table Name : SpexPiston
# Column Name : CompressionHeight_IN
cursor.execute('SELECT CompressionHeight_IN FROM SpexPiston')
compression_height_by_inch = []
for i in cursor:
    # print(i)
    compression_height_by_inch.append(i)
print("[#]Compression Height By Inch:")
# [25000:25003] : Range of List needed by index
print("     ", compression_height_by_inch[25000:25003])
# endregion  <<<<===========================[Compression Height By Inch]===========================>>>>

# region  <<<<============================[Dome Height]============================>>>>
# Table Name : SpexPiston
# Column Name : DomeHeight
cursor.execute('SELECT DomeHeight FROM SpexPiston')
dome_height = []
for i in cursor:
    # print(i)
    dome_height.append(i)
print("[#]Dome Height:")
# [25000:25003] : Range of List needed by index
print("     ", dome_height[25000:25003])
# endregion  <<<<===========================[Dome Height]===========================>>>>

# region  <<<<============================[Deck Thickness]============================>>>>
# Table Name : SpexPiston
# Column Name : DeckThickness
cursor.execute('SELECT DeckThickness FROM SpexPiston')
deck_thickness = []
for i in cursor:
    # print(i)
    deck_thickness.append(i)
print("[#]Deck Thickness:")
# [25000:25003] : Range of List needed by index
print("     ", deck_thickness[25000:25003])
# endregion  <<<<===========================[Deck Thickness]===========================>>>>

# region  <<<<============================[Target Weight]============================>>>>
# Table Name : SpexPiston
# Column Name : TargetWeight
cursor.execute('SELECT TargetWeight FROM SpexPiston')
target_weight = []
for i in cursor:
    # print(i)
    target_weight.append(i)
print("[#]Target Weight:")
# [25000:25003] : Range of List needed by index
print("     ", target_weight[25000:25003])
# endregion  <<<<===========================[Target Weight]===========================>>>>

# region  <<<<============================[CAD Files]============================>>>>
# Table Name : SpexPiston
# Column Name : CADFiles
cursor.execute('SELECT CADFiles FROM SpexPiston')
CAD_files = []
for i in cursor:
    # print(i)
    CAD_files.append(i)
print("[#]CAD Files:")
# [25000:25003] : Range of List needed by index
print("     ", CAD_files[25000:25003])
# endregion  <<<<===========================[CAD Files]===========================>>>>

# region  <<<<============================[Reference Part Number]============================>>>>
# Table Name : SpexPiston
# Column Name : ReferencePartNumber
cursor.execute('SELECT ReferencePartNumber FROM SpexPiston')
reference_part_number = []
for i in cursor:
    # print(i)
    reference_part_number.append(i)
print("[#]Reference Part Number:")
# [25000:25003] : Range of List needed by index
print("     ", reference_part_number[25000:25003])
# endregion  <<<<===========================[Reference Part Number]===========================>>>>

# region  <<<<============================[Pilot Bore Diameter]============================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBoreDiameter
cursor.execute('SELECT PilotBoreDiameter FROM SpexPiston_PinBore')
pilot_bore_diameter = []
for i in cursor:
    # print(i)
    pilot_bore_diameter.append(i)
print("[#]Pilot Bore Diameter:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_diameter[25000:25003])
# endregion  <<<<===========================[Pilot Bore Diameter]===========================>>>>

# region  <<<<===================[Pilot Bore Pilot To Pin]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBorePilotToPin
cursor.execute('SELECT PilotBorePilotToPin FROM SpexPiston_PinBore')
pilot_bore_pilot_to_pin = []
for i in cursor:
    # print(i)
    pilot_bore_pilot_to_pin.append(i)
print("[#]Pilot Bore Pilot To Pin:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_pilot_to_pin[25000:25003])
# endregion  <<<<===================[Pilot Bore Pilot To Pin]===================>>>>

# region  <<<<===================[Pilot Bore Depth To Dome]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBoreDepthToDome
cursor.execute('SELECT PilotBoreDepthToDome FROM SpexPiston_PinBore')
pilot_bore_depth_to_dome = []
for i in cursor:
    # print(i)
    pilot_bore_depth_to_dome.append(i)
print("[#]Pilot Bore Depth To Dome:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_depth_to_dome[25000:25003])
# endregion  <<<<===================[Pilot Bore Depth To Dome]===================>>>>

# region  <<<<===================[Pilot Bore Pilot To Belt]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBorePilotToBelt
cursor.execute('SELECT PilotBorePilotToBelt FROM SpexPiston_PinBore')
pilot_bore_pilot_to_Belt = []
for i in cursor:
    # print(i)
    pilot_bore_pilot_to_Belt.append(i)
print("[#]Pilot Bore Pilot To Belt:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_pilot_to_Belt[25000:25003])
# endregion  <<<<===================[Pilot Bore Pilot To Belt]===================>>>>

# region  <<<<===================[Pilot Bore Depth To Deck]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBoreDepthToDeck
cursor.execute('SELECT PilotBoreDepthToDeck FROM SpexPiston_PinBore')
pilot_bore_depth_to_deck = []
for i in cursor:
    # print(i)
    pilot_bore_depth_to_deck.append(i)
print("[#]Pilot Bore Depth To Deck:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_depth_to_deck[25000:25003])
# endregion  <<<<===================[Pilot Bore Depth To Deck]===================>>>>

# region  <<<<===================[Pilot Bore Pilot To TE]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBorePilotToTE
cursor.execute('SELECT PilotBorePilotToTE FROM SpexPiston_PinBore')
pilot_bore_pilot_to_TE = []
for i in cursor:
    # print(i)
    pilot_bore_pilot_to_TE.append(i)
print("[#]Pilot Bore Pilot To TE:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_pilot_to_TE[25000:25003])
# endregion  <<<<===================[Pilot Bore Pilot To TE]===================>>>>

# region  <<<<===================[Pilot Bore Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PilotBoreNotes
cursor.execute('SELECT PilotBoreNotes FROM SpexPiston_PinBore')
pilot_bore_notes = []
for i in cursor:
    # print(i)
    pilot_bore_notes.append(i)
print("[#]Pilot Bore Notes:")
# [25000:25003] : Range of List needed by index
print("     ", pilot_bore_notes[25000:25003])
# endregion  <<<<===================[Pilot Bore Notes]===================>>>>

# region  <<<<===================[Piston Pin Bore Diameter By Inch]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreDiameterIN
cursor.execute('SELECT PistPinBoreDiameterIN FROM SpexPiston_PinBore')
piston_pin_bore_diameter_by_inch = []
for i in cursor:
    # print(i)
    piston_pin_bore_diameter_by_inch.append(i)
    PistPinBoreDiameterIN = i
# just for now
# THAT'S HOW YOU GET PistonID TO USE IT TO LOOK UP THE PISTON NUMBER FOR SPECIFIC FILTER
# cursor.execute('SELECT PistPinBoreDiameterIN,PistonID FROM SpexPiston_PinBore WHERE PistPinBoreDiameterIN = 0.4710')
# That's how you get two columns from different tables
# cursor.execute(
#     "SELECT PistPinBoreDiameterIN,Piston FROM SpexPiston_PinBore, SpexPiston")
# Needs to Delete below later
# cursor.execute(
#     "SELECT PistPinBoreDiameterIN FROM SpexPiston_PinBore")
# piston_pin_bore_diameter_by_inch = []
# for i in cursor:
#     print(i)
#     piston_pin_bore_diameter_by_inch.append(i)
#     # cursor.execute('SELECT Piston FROM SpexPiston WHERE ' + PistPinBoreDiameterIN + ' = ?', 0)
#     # for i in cursor:
#     #     print(i)
# cursor.execute(
#     "SELECT Piston FROM SpexPiston")
# piston_pin_bore_diameter_by_inch = []
# for i in cursor:
#     print(i)

print("[#]Piston Pin Bore Diameter By Inch:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_diameter_by_inch[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Diameter By Inch]===================>>>>

# region  <<<<===================[Piston Pin Bore Diameter By Millimeter]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreDiameterMM
cursor.execute('SELECT PistPinBoreDiameterMM FROM SpexPiston_PinBore')
piston_pin_bore_diameter_by_millimeter = []
for i in cursor:
    # print(i)
    piston_pin_bore_diameter_by_millimeter.append(i)
print("[#]Piston Pin Bore Diameter By Millimeter:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_diameter_by_millimeter[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Diameter By Millimeter]===================>>>>

# region  <<<<===================[Piston Pin Bore Offset Amount]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreOffsetAmount
cursor.execute('SELECT PistPinBoreOffsetAmount FROM SpexPiston_PinBore')
piston_pin_bore_offset_amount = []
for i in cursor:
    # print(i)
    piston_pin_bore_offset_amount.append(i)
print("[#]Piston Pin Bore Offset Amount:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_offset_amount[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Offset Amount]===================>>>>

# region  <<<<===================[Piston Pin Bore Offset Direction]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreOffsetDirection
cursor.execute('SELECT PistPinBoreOffsetDirection FROM SpexPiston_PinBore')
piston_pin_bore_offset_direction = []
for i in cursor:
    # print(i)
    piston_pin_bore_offset_direction.append(i)
print("[#]Piston Pin Bore Offset Direction:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_offset_direction[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Offset Direction]===================>>>>

# region  <<<<===================[Piston Pin Bore Bump Identification]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreSkirtIdent
cursor.execute('SELECT PistPinBoreSkirtIdent FROM SpexPiston_PinBore')
piston_pin_bore_bump_Identification = []
for i in cursor:
    # print(i)
    piston_pin_bore_bump_Identification.append(i)
print("[#]Piston Pin Bore Bump Identification:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_bump_Identification[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Bump Identification]===================>>>>

# region  <<<<===================[Piston Pin Bore Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : PistPinBoreNotes
cursor.execute('SELECT PistPinBoreNotes FROM SpexPiston_PinBore')
piston_pin_bore_notes = []
for i in cursor:
    # print(i)
    piston_pin_bore_notes.append(i)
print("[#]Piston Pin Bore Notes:")
# [25000:25003] : Range of List needed by index
print("     ", piston_pin_bore_notes[25000:25003])
# endregion  <<<<===================[Piston Pin Bore Notes]===================>>>>

# region  <<<<===================[RetClip Grv Diameter]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipGrvDiameter
cursor.execute('SELECT RetClipGrvDiameter FROM SpexPiston_PinBore')
ret_clip_groove_diameter = []
for i in cursor:
    # print(i)
    ret_clip_groove_diameter.append(i)
print("[#]RetClip Grv Diameter:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_groove_diameter[25000:25003])
# endregion  <<<<===================[RetClip Grv Diameter]===================>>>>

# region  <<<<===================[RetClip Grv Width]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipGrvWidth
cursor.execute('SELECT RetClipGrvWidth FROM SpexPiston_PinBore')
ret_clip_groove_width = []
for i in cursor:
    # print(i)
    ret_clip_groove_width.append(i)
print("[#]RetClip Grv Width:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_groove_width[25000:25003])
# endregion  <<<<===================[RetClip Grv Width]===================>>>>

# region  <<<<===================[RetClip Grv ID Spacing]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipGrvInnerDiameterSpace
cursor.execute('SELECT RetClipGrvInnerDiameterSpace FROM SpexPiston_PinBore')
ret_clip_groove_id_spacing = []
for i in cursor:
    # print(i)
    ret_clip_groove_id_spacing.append(i)
print("[#]RetClip Grv ID Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_groove_id_spacing[25000:25003])
# endregion  <<<<===================[RetClip Grv ID Spacing]===================>>>>

# region  <<<<===================[RetClip Grv Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotes
cursor.execute('SELECT RetClipNotes FROM SpexPiston_PinBore')
ret_clip_groove_notes = []
for i in cursor:
    # print(i)
    ret_clip_groove_notes.append(i)
print("[#]RetClip Grv Notes:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_groove_notes[25000:25003])
# endregion  <<<<===================[RetClip Grv Notes]===================>>>>

# region  <<<<===================[CFren Grv Diameter]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : CFrenGrvDiameter
cursor.execute('SELECT CFrenGrvDiameter FROM SpexPiston_PinBore')
CFren_groove_diameter = []
for i in cursor:
    # print(i)
    CFren_groove_diameter.append(i)
print("[#]CFren Grv Diameter:")
# [25000:25003] : Range of List needed by index
print("     ", CFren_groove_diameter[25000:25003])
# endregion  <<<<===================[CFren Grv Diameter]===================>>>>

# region  <<<<===================[CFren Grv ID Spacing]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : CFrenGrvInnerDiameterSpace
cursor.execute('SELECT CFrenGrvInnerDiameterSpace FROM SpexPiston_PinBore')
CFren_groove_id_spacing = []
for i in cursor:
    # print(i)
    CFren_groove_id_spacing.append(i)
print("[#]CFren Grv ID Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", CFren_groove_id_spacing[25000:25003])
# endregion  <<<<===================[CFren Grv ID Spacing]===================>>>>

# region  <<<<===================[CFren Grv Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : CFrenNotes
cursor.execute('SELECT CFrenNotes FROM SpexPiston_PinBore')
CFren_groove_notes = []
for i in cursor:
    # print(i)
    CFren_groove_notes.append(i)
print("[#]CFren Grv Notes:")
# [25000:25003] : Range of List needed by index
print("     ", CFren_groove_notes[25000:25003])
# endregion  <<<<===================[CFren Grv Notes]===================>>>>

# region  <<<<===================[Semi CFren Grv Width]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : SemiCFrenGrvWidth
cursor.execute('SELECT SemiCFrenGrvWidth FROM SpexPiston_PinBore')
semi_CFren_groove_width = []
for i in cursor:
    # print(i)
    semi_CFren_groove_width.append(i)
print("[#]Semi CFren Grv Width:")
# [25000:25003] : Range of List needed by index
print("     ", semi_CFren_groove_width[25000:25003])
# endregion  <<<<===================[Semi CFren Grv Width]===================>>>>

# region  <<<<===================[Semi CFren Grv Depth]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : SemiCFrenGrvDepth
cursor.execute('SELECT SemiCFrenGrvDepth FROM SpexPiston_PinBore')
semi_CFren_groove_depth = []
for i in cursor:
    # print(i)
    semi_CFren_groove_depth.append(i)
print("[#]Semi CFren Grv Depth:")
# [25000:25003] : Range of List needed by index
print("     ", semi_CFren_groove_depth[25000:25003])
# endregion  <<<<===================[Semi CFren Grv Depth]===================>>>>

# region  <<<<===================[Semi CFren Grv ID Spacing]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : SemiCFrenGrvInnerDiameterSpace
cursor.execute('SELECT SemiCFrenGrvInnerDiameterSpace FROM SpexPiston_PinBore')
semi_CFren_groove_id_spacing = []
for i in cursor:
    # print(i)
    semi_CFren_groove_id_spacing.append(i)
print("[#]Semi CFren Grv ID Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", semi_CFren_groove_id_spacing[25000:25003])
# endregion  <<<<===================[Semi CFren Grv ID Spacing]===================>>>>

# region  <<<<===================[Semi CFren Grv Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : SemiCFrenGrvNotes
cursor.execute('SELECT SemiCFrenGrvNotes FROM SpexPiston_PinBore')
semi_CFren_groove_notes = []
for i in cursor:
    # print(i)
    semi_CFren_groove_notes.append(i)
print("[#]Semi CFren Grv Notes:")
# [25000:25003] : Range of List needed by index
print("     ", semi_CFren_groove_notes[25000:25003])
# endregion  <<<<===================[Semi CFren Grv Notes]===================>>>>

# region  <<<<===================[RetClip Notch Diameter Depth]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotchDiameterDepth
cursor.execute('SELECT RetClipNotchDiameterDepth FROM SpexPiston_PinBore')
ret_clip_notch_diameter_depth = []
for i in cursor:
    # print(i)
    ret_clip_notch_diameter_depth.append(i)
print("[#]RetClip Notch Diameter Depth:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_notch_diameter_depth[25000:25003])
# endregion  <<<<===================[RetClip Notch Diameter Depth]===================>>>>

# region  <<<<===================[RetClip Notch 1st Location Angle]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotchLocAngle01
cursor.execute('SELECT RetClipNotchLocAngle01 FROM SpexPiston_PinBore')
ret_clip_notch_first_location_angle = []
for i in cursor:
    # print(i)
    ret_clip_notch_first_location_angle.append(i)
print("[#]RetClip Notch 1st Location Angle:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_notch_first_location_angle[25000:25003])
# endregion  <<<<===================[RetClip Notch 1st Location Angle]===================>>>>

# region  <<<<===================[RetClip Notch 2nd Location Angle]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotchLocAngle02
cursor.execute('SELECT RetClipNotchLocAngle02 FROM SpexPiston_PinBore')
ret_clip_notch_second_location_angle = []
for i in cursor:
    # print(i)
    ret_clip_notch_second_location_angle.append(i)
print("[#]RetClip Notch 2nd Location Angle:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_notch_second_location_angle[25000:25003])
# endregion  <<<<===================[RetClip Notch 2nd Location Angle]===================>>>>

# region  <<<<===================[RetClip Notch Depth Grv]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotchDepthGrv
cursor.execute('SELECT RetClipNotchDepthGrv FROM SpexPiston_PinBore')
ret_clip_notch_depth_grv = []
for i in cursor:
    # print(i)
    ret_clip_notch_depth_grv.append(i)
print("[#]RetClip Notch Depth Grv:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_notch_depth_grv[25000:25003])
# endregion  <<<<===================[RetClip Notch Depth Grv]===================>>>>

# region  <<<<===================[RetClip Notch Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : RetClipNotchNotes
cursor.execute('SELECT RetClipNotchNotes FROM SpexPiston_PinBore')
ret_clip_notch_notes = []
for i in cursor:
    # print(i)
    ret_clip_notch_notes.append(i)
print("[#]RetClip Notch Notes:")
# [25000:25003] : Range of List needed by index
print("     ", ret_clip_notch_notes[25000:25003])
# endregion  <<<<===================[RetClip Notch Notes]===================>>>>

# region  <<<<===================[Horizontal Slots Diameter Depth]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsDiameterDepth
cursor.execute('SELECT HorizPistonPinSlotsDiameterDepth FROM SpexPiston_PinBore')
horizontal_slots_diameter_depth = []
for i in cursor:
    # print(i)
    horizontal_slots_diameter_depth.append(i)
print("[#]Horizontal Slots Diameter Depth:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_diameter_depth[25000:25003])
# endregion  <<<<===================[Horizontal Slots Diameter Depth]===================>>>>

# region  <<<<===================[Horizontal Slots Width]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsWidth
cursor.execute('SELECT HorizPistonPinSlotsWidth FROM SpexPiston_PinBore')
horizontal_slots_depth = []
for i in cursor:
    # print(i)
    horizontal_slots_depth.append(i)
print("[#]Horizontal Slots Width:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_depth[25000:25003])
# endregion  <<<<===================[Horizontal Slots Width]===================>>>>

# region  <<<<===================[Distance Between Horizontal Slots]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsDistBetSlots
cursor.execute('SELECT HorizPistonPinSlotsDistBetSlots FROM SpexPiston_PinBore')
distance_between_horizontal_slots_depth = []
for i in cursor:
    # print(i)
    distance_between_horizontal_slots_depth.append(i)
print("[#]Distance Between Horizontal Slots:")
# [25000:25003] : Range of List needed by index
print("     ", distance_between_horizontal_slots_depth[25000:25003])
# endregion  <<<<===================[Distance Between Horizontal Slots]===================>>>>

# region  <<<<===================[Horizontal Slots Through Boss Status]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsThruBoss
cursor.execute('SELECT HorizPistonPinSlotsThruBoss FROM SpexPiston_PinBore')
horizontal_slots_through_Boss_status = []
for i in cursor:
    # print(i)
    horizontal_slots_through_Boss_status.append(i)
print("[#]Horizontal Slots Through Boss Status:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_through_Boss_status[25000:25003])
# endregion  <<<<===================[Horizontal Slots Through Boss Status]===================>>>>

# region  <<<<===================[Horizontal Slots OD Spacing]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsOuterDiameterSpacing
cursor.execute('SELECT HorizPistonPinSlotsOuterDiameterSpacing FROM SpexPiston_PinBore')
horizontal_slots_OD_spacing = []
for i in cursor:
    # print(i)
    horizontal_slots_OD_spacing.append(i)
print("[#]Horizontal Slots OD Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_OD_spacing[25000:25003])
# endregion  <<<<===================[Horizontal Slots OD Spacing]===================>>>>

# region  <<<<===================[Horizontal Slots Arc Diameter]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsArcDiameter
cursor.execute('SELECT HorizPistonPinSlotsArcDiameter FROM SpexPiston_PinBore')
horizontal_slots_arc_diameter = []
for i in cursor:
    # print(i)
    horizontal_slots_arc_diameter.append(i)
print("[#]Horizontal Slots OD Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_arc_diameter[25000:25003])
# endregion  <<<<===================[Horizontal Slots Arc Diameter]===================>>>>

# region  <<<<===================[Horizontal Slots Notes]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : HorizPistonPinSlotsNotes
cursor.execute('SELECT HorizPistonPinSlotsNotes FROM SpexPiston_PinBore')
horizontal_slots_notes = []
for i in cursor:
    # print(i)
    horizontal_slots_notes.append(i)
print("[#]Horizontal Slots Notes:")
# [25000:25003] : Range of List needed by index
print("     ", horizontal_slots_notes[25000:25003])
# endregion  <<<<===================[Horizontal Slots Notes]===================>>>>

# region  <<<<===================[Legacy PinBore Comments]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : LegacyPinBoreComments
cursor.execute('SELECT LegacyPinBoreComments FROM SpexPiston_PinBore')
legacy_pin_bore_comments = []
for i in cursor:
    # print(i)
    legacy_pin_bore_comments.append(i)
print("[#]Legacy PinBore Comments:")
# [25000:25003] : Range of List needed by index
print("     ", legacy_pin_bore_comments[25000:25003])
# endregion  <<<<===================[Legacy PinBore Comments]===================>>>>

# region  <<<<===================[Legacy PinBore Gas Comments]===================>>>>
# Table Name : SpexPiston_PinBore
# Column Name : LegacyPinBoreGasComments
cursor.execute('SELECT LegacyPinBoreGasComments FROM SpexPiston_PinBore')
legacy_pin_bore_gas_comments = []
for i in cursor:
    # print(i)
    legacy_pin_bore_gas_comments.append(i)
print("[#]Legacy PinBore Gas Comments:")
# [25000:25003] : Range of List needed by index
print("     ", legacy_pin_bore_gas_comments[25000:25003])
# endregion  <<<<================[Legacy PinBore Gas Comments]===================>>>>

# region  <<<<===================[Piston Overall Length Intake]===================>>>>
# Table Name : SpexPiston_SemiFinishTurn
# Column Name : OverallLengthInt
cursor.execute('SELECT OverallLengthInt FROM SpexPiston_SemiFinishTurn')
piston_overall_length_intake = []
for i in cursor:
    # print(i)
    piston_overall_length_intake.append(i)
print("[#]Piston Overall Length Intake:")
# [25000:25003] : Range of List needed by index
print("     ", piston_overall_length_intake[25000:25003])
# endregion  <<<<===================[Piston Overall Length Intake]===================>>>>

# region  <<<<===================[Piston Overall Length Exhaust]===================>>>>
# Table Name : SpexPiston_SemiFinishTurn
# Column Name : OverallLengthExh
cursor.execute('SELECT OverallLengthExh FROM SpexPiston_SemiFinishTurn')
piston_overall_length_exhaust = []
for i in cursor:
    # print(i)
    piston_overall_length_exhaust.append(i)
print("[#]Piston Overall Length Exhaust:")
# [25000:25003] : Range of List needed by index
print("     ", piston_overall_length_exhaust[25000:25003])
# endregion  <<<<===================[Piston Overall Length Exhaust]===================>>>>

# region  <<<<===================[Skirt Mill Intake OverallLength From TE]===================>>>>
# Table Name : SpexPiston_Milling
# Column Name : SkirtMillIntakeOverallLengthFromTE
cursor.execute('SELECT SkirtMillIntakeOverallLengthFromTE FROM SpexPiston_Milling')
skirt_mill_intake_overall_length_from_TE = []
for i in cursor:
    # print(i)
    skirt_mill_intake_overall_length_from_TE.append(i)
print("[#]Skirt Mill Intake OverallLength From TE:")
# [25000:25003] : Range of List needed by index
print("     ", skirt_mill_intake_overall_length_from_TE[25000:25003])
# endregion  <<<<===================[Skirt Mill Intake OverallLength From TE]===================>>>>

# region  <<<<===================[Skirt Mill Exhaust OverallLength From TE]===================>>>>
# Table Name : SpexPiston_Milling
# Column Name : SkirtMillExhaustOverallLengthFromTE
cursor.execute('SELECT SkirtMillExhaustOverallLengthFromTE FROM SpexPiston_Milling')
skirt_mill_Exhaust_overall_length_from_TE = []
for i in cursor:
    # print(i)
    skirt_mill_Exhaust_overall_length_from_TE.append(i)
print("[#]Skirt Mill Exhaust OverallLength From TE:")
# [25000:25003] : Range of List needed by index
print("     ", skirt_mill_Exhaust_overall_length_from_TE[25000:25003])
# endregion  <<<<===================[Skirt Mill Exhaust OverallLength From TE]===================>>>>

# region  <<<<===================[Pressure Fed Oil Hole Type]===================>>>>
# Table Name : SpexPiston_GasPortsPistonPinOiling
# Column Name : PressureFedOilHoleType
cursor.execute('SELECT PressureFedOilHoleType FROM SpexPiston_GasPortsPistonPinOiling')
pressure_fed_oil_hole_type = []
for i in cursor:
    # print(i)
    pressure_fed_oil_hole_type.append(i)
print("[#]Pressure Fed Oil Hole Type:")
# [25000:25003] : Range of List needed by index
print("     ", pressure_fed_oil_hole_type[25000:25003])
# endregion  <<<<===================[Pressure Fed Oil Hole Type]===================>>>>

# region  <<<<===================[Pressure Fed Oil Hole Notes]===================>>>>
# Table Name : SpexPiston_GasPortsPistonPinOiling
# Column Name : PressureFedOilHoleNotes
cursor.execute('SELECT PressureFedOilHoleNotes FROM SpexPiston_GasPortsPistonPinOiling')
pressure_fed_oil_hole_notes = []
for i in cursor:
    # print(i)
    pressure_fed_oil_hole_notes.append(i)
print("[#]Pressure Fed Oil Hole Notes:")
# [25000:25003] : Range of List needed by index
print("     ", pressure_fed_oil_hole_notes[25000:25003])
# endregion  <<<<===================[Pressure Fed Oil Hole Notes]===================>>>>

# region  <<<<===================[Double Oil Holes Slots ID Spacing]===================>>>>
# Table Name : SpexPiston_GasPortsPistonPinOiling
# Column Name : PressureFedOilHoleInnerDiameterSpace
cursor.execute('SELECT PressureFedOilHoleInnerDiameterSpace FROM SpexPiston_GasPortsPistonPinOiling')
double_oil_holes_slots_id_spacing = []
for i in cursor:
    # print(i)
    double_oil_holes_slots_id_spacing.append(i)
print("[#]Double Oil Holes Slots ID Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", double_oil_holes_slots_id_spacing[25000:25003])
# endregion  <<<<===================[Double Oil Holes Slots ID Spacing]===================>>>>


#          <<<<<<<<------------------------------------------------------------------------>>>>>>>>>>
#      <<<<<<<<--------------------------------------------------------------------------------->>>>>>>>>>
# <<<<<<<<------------------------------------------------------------------------------------------->>>>>>>>>>
# endregion  <<<<===========================[Spex Information from DataBase]===========================>>>>

print()
# region  <<<<===================[To Get PistonID By Giving Job Number]===================>>>>
# Needs PistonID to get any Value you need like PilotToPin, Offset, OilGrvDiameter...Etc
# define variable to have the job number
job_number = 'WIS-10128'
# Set PistonID by having the Job Number
# That's how to use variable inside SQL query
cursor.execute("SELECT PistonID FROM SpexPiston WHERE Piston = ?", job_number)
# (fetchone()): USED TO SHOW JUST THE VALUE OF THE DATA WE LOOKING FOR
for i in cursor.fetchone():
    PistonID = i
print("[#]PistonID FOR " + job_number + ":")
print("     ", PistonID)
# endregion  <<<<=================[To Get PistonID By Giving Job Number]===================>>>>

# region  <<<<===================[Template to Get Specific Value For Specific Job]===================>>>>
# Change Table & Column Names according to value you looking for
# Table Name & Column Name
# Change 'PistonID' to check different job number direct or it'll change according to the job number above
# Ex: use PistonID "25002" which indicate job number "WD-07971"
cursor.execute('SELECT PistPinBoreDiameterIN FROM '
               'SpexPiston_PinBore WHERE PistonID = ?', PistonID)
specific_value_for_specific_job = []
for i in cursor.fetchone():
    # print(i)
    specific_value_for_specific_job = i
print("[#]Specific_Value_For_Specific_Job_for " + job_number + ":")
# [25000:25003] : Range of List needed by index
print("     ", specific_value_for_specific_job)
print(type(specific_value_for_specific_job))
# if (type(specific_value_for_specific_job) == str):
#     print("Type is ")
# substr = '0.076'
# # for line in specific_value_for_specific_job.rstrip(' '):
# index = specific_value_for_specific_job.find(substr)
# if index != -1:
#     print("index is ", index)
#     print(specific_value_for_specific_job[index])
# else:
#     print("not found")

# endregion  <<<<=================[Template to Get Specific Value For Specific Job]===================>>>>

# region  <<<<===================[To Get Forging Number By Having The Job Number]===================>>>>
# First,needs to get 'ForgeSpecID' from 'SpexPiston' table by passing the 'PistonID'(that we got before from job number)
cursor.execute('SELECT ForgeSpecID FROM SpexPiston WHERE PistonID = ?', PistonID)
for i in cursor.fetchone():
    ForgeSpecID = i
print("[#]ForgeSpecID For Job " + job_number + ":")
print("     ", ForgeSpecID)
# Then, needs to get 'ForgeItemID'(Forging Number) from 'SpexForge' table by passing the 'ForgeSpecID'
cursor.execute('SELECT ForgeItemID FROM SpexForge WHERE ForgeSpecID = ?', ForgeSpecID)
for i in cursor.fetchone():
    forging_number = i
print("[#]Forging Number For Job " + job_number + ":")
print("     ", forging_number)
# endregion  <<<<===============[To Get Forging Number By Having The Job Number]================>>>>

# region  <<<<===================[To Filter Forging Number to Use It to Find Probe Program]===================>>>>
letter = 0
forging_number_for_probe_program = "F"
# forging_number[-1]: IS THE LAST DIGIT OF STRING
while (letter != "X" and letter != "x" and letter != "Z" and letter != "z" and letter != forging_number[-1]):
    for letter in forging_number[1:]:
        print(letter)
        forging_number_for_probe_program = forging_number_for_probe_program + letter
        if (letter == "X" or letter == "x" or letter == "Z" or letter == "z"):
            break
print(forging_number_for_probe_program)
# endregion  <<<<=================[To Filter Forging Number to Use It to Find Probe Program]==================>>>>

print()
# region  <<<<============================[Forging Information from DataBase]============================>>>>
# <<<<<<<<------------------------------------------------------------------------------------------->>>>>>>>>>
#      <<<<<<<<--------------------------------------------------------------------------------->>>>>>>>>>
#          <<<<<<<<------------------------------------------------------------------------>>>>>>>>>>

print("<<<<============================[Forging Information from DataBase]============================>>>>")
# region  <<<<============================[Forge Spec ID]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeSpecID
cursor.execute('SELECT ForgeSpecID FROM SpexForge')
forge_Spec_id = []
for i in cursor:
    # print(i)
    forge_Spec_id.append(i)
print("[#]Forge Spec ID:")
# [25000:25003] : Range of List needed by index
print("     ", forge_Spec_id[200:203])
# endregion  <<<<===========================[Forge Spec ID]===========================>>>>

# region  <<<<============================[Forge Item ID (Forging Number)]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeItemID
cursor.execute('SELECT ForgeItemID FROM SpexForge')
forge_item_id = []
for i in cursor:
    # print(i)
    forge_item_id.append(i)
print("[#]Forge Item ID (Forging Number):")
# [25000:25003] : Range of List needed by index
print("     ", forge_item_id[200:203])
# endregion  <<<<==========================[Forge Item ID (Forging Number)]===========================>>>>

# region  <<<<============================[Forge Rev Level]============================>>>>
# Table Name : SpexForge
# Column Name : RevLevel
cursor.execute('SELECT RevLevel FROM SpexForge')
forge_rev_level = []
for i in cursor:
    # print(i)
    forge_rev_level.append(i)
print("[#]Forge Rev Level:")
# [25000:25003] : Range of List needed by index
print("     ", forge_rev_level[200:203])
# endregion  <<<<==========================[Forge Rev Level]===========================>>>>

# region  <<<<============================[Forge Ref Length]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeRefLength
cursor.execute('SELECT ForgeRefLength FROM SpexForge')
forge_ref_length = []
for i in cursor:
    # print(i)
    forge_ref_length.append(i)
print("[#]Forge Ref Length:")
# [25000:25003] : Range of List needed by index
print("     ", forge_ref_length[200:203])
# endregion  <<<<==========================[Forge Ref Length]===========================>>>>

# region  <<<<============================[Forge OD At Rougher]============================>>>>
# Table Name : SpexForge
# Column Name : ODAtRougher
cursor.execute('SELECT ODAtRougher FROM SpexForge')
forge_OD_at_rougher = []
for i in cursor:
    # print(i)
    forge_OD_at_rougher.append(i)
print("[#]Forge OD At Rougher:")
# [25000:25003] : Range of List needed by index
print("     ", forge_OD_at_rougher[200:203])
# endregion  <<<<==========================[Forge OD At Rougher]===========================>>>>

# region  <<<<============================[Forge Boss Outside Spacing]============================>>>>
# Table Name : SpexForge
# Column Name : BossOutsdSpace
cursor.execute('SELECT BossOutsdSpace FROM SpexForge')
forge_boss_outside_spacing = []
for i in cursor:
    # print(i)
    forge_boss_outside_spacing.append(i)
print("[#]Forge Boss Outside Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", forge_boss_outside_spacing[200:203])
# endregion  <<<<==========================[Forge Boss Outside Spacing]===========================>>>>

# region  <<<<============================[Forge Boss Inside Spacing]============================>>>>
# Table Name : SpexForge
# Column Name : BossInsdSpace
cursor.execute('SELECT BossInsdSpace FROM SpexForge')
forge_boss_inside_spacing = []
for i in cursor:
    # print(i)
    forge_boss_inside_spacing.append(i)
print("[#]Forge Boss Inside Spacing:")
# [25000:25003] : Range of List needed by index
print("     ", forge_boss_inside_spacing[200:203])
# endregion  <<<<==========================[Forge Boss Inside Spacing]===========================>>>>

# region  <<<<============================[Forge Style]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeStyle
cursor.execute('SELECT ForgeStyle FROM SpexForge')
forge_style = []
for i in cursor:
    # print(i)
    forge_style.append(i)
print("[#]Forge Style:")
# [25000:25003] : Range of List needed by index
print("     ", forge_style[200:203])
# endregion  <<<<==========================[Forge Style]===========================>>>>

# region  <<<<============================[Forge Ring Belt Height Outside]============================>>>>
# Table Name : SpexForge
# Column Name : RingBeltHtOutsd
cursor.execute('SELECT RingBeltHtOutsd FROM SpexForge')
forge_ring_belt_Height_Outside = []
for i in cursor:
    # print(i)
    forge_ring_belt_Height_Outside.append(i)
print("[#]Forge Ring Belt Height Outside:")
# [25000:25003] : Range of List needed by index
print("     ", forge_ring_belt_Height_Outside[200:203])
# Note: (X) Outsd_Ring_Belt_Ht in EMSS
# endregion  <<<<==========================[Forge Ring Belt Height Outside]===========================>>>>

# region  <<<<============================[Forge Ring Belt Inside Diameter]============================>>>>
# Table Name : SpexForge
# Column Name : RingBeltID
cursor.execute('SELECT RingBeltID FROM SpexForge')
forge_ring_belt_id = []
for i in cursor:
    # print(i)
    forge_ring_belt_id.append(i)
print("[#]Forge Ring Belt Inside Diameter:")
# [25000:25003] : Range of List needed by index
print("     ", forge_ring_belt_id[200:203])
# Note: (S) Ring Belt I.D. in EMSS
# endregion  <<<<==========================[Forge Ring Belt Inside Diameter]===========================>>>>

# region  <<<<============================[Forge Min Dome At Rougher]============================>>>>
# Table Name : SpexForge
# Column Name : MinDomeAtRougher
cursor.execute('SELECT MinDomeAtRougher FROM SpexForge')
forge_min_dome_at_rougher = []
for i in cursor:
    # print(i)
    forge_min_dome_at_rougher.append(i)
print("[#]Forge Min Dome At Rougher:")
# [2000:2003] : Range of List needed by index
print("     ", forge_min_dome_at_rougher[200:203])
# endregion  <<<<==========================[Forge Min Dome At Rougher]===========================>>>>

# region  <<<<============================[Forge Min Dome At Forge]============================>>>>
# Table Name : SpexForge
# Column Name : MinDomeAtForge
cursor.execute('SELECT MinDomeAtForge FROM SpexForge')
forge_min_dome_at_forge = []
for i in cursor:
    # print(i)
    forge_min_dome_at_forge.append(i)
print("[#]Forge Min Dome At Forge:")
# [2000:2003] : Range of List needed by index
print("     ", forge_min_dome_at_forge[200:203])
# endregion  <<<<==========================[Forge Min Dome At Forge]===========================>>>>

# region  <<<<============================[Forge Probe Program]============================>>>>
# Table Name : SpexForge
# Column Name : ProbeProgram
cursor.execute('SELECT ProbeProgram FROM SpexForge')
forge_probe_program = []
for i in cursor:
    # print(i)
    forge_probe_program.append(i)
print("[#]Forge ProbeProgram:")
# [2000:2003] : Range of List needed by index
print("     ", forge_probe_program[200:203])
# endregion  <<<<==========================[Forge Probe Program]===========================>>>>

# region  <<<<============================[Forge Active Status]============================>>>>
# Table Name : SpexForge
# Column Name : Active
cursor.execute('SELECT Active FROM SpexForge')
forge_active_status = []
for i in cursor:
    # print(i)
    forge_active_status.append(i)
print("[#]Forge Active Status:")
# [2000:2003] : Range of List needed by index
print("     ", forge_active_status[200:203])
# endregion  <<<<==========================[Forge Active Status]===========================>>>>

# region  <<<<============================[Engine Type]============================>>>>
# Table Name : SpexForge
# Column Name : EngineType
cursor.execute('SELECT EngineType FROM SpexForge')
engine_type = []
for i in cursor:
    # print(i)
    engine_type.append(i)
print("[#]Engine Type:")
# [2000:2003] : Range of List needed by index
print("     ", engine_type[200:203])
# endregion  <<<<==========================[Engine Type]===========================>>>>

# region  <<<<============================[Forging Material]============================>>>>
# Table Name : SpexForge
# Column Name : Material
cursor.execute('SELECT Material FROM SpexForge')
forging_material = []
for i in cursor:
    # print(i)
    forging_material.append(i)
print("[#]Forging Material:")
# [2000:2003] : Range of List needed by index
print("     ", forging_material[200:203])
# endregion  <<<<==========================[Forging Material]===========================>>>>

# region  <<<<============================[Forging ERP Item ID]============================>>>>
# Table Name : SpexForge
# Column Name : ERP_ItemID
cursor.execute('SELECT ERP_ItemID FROM SpexForge')
forging_ERP_item_id = []
for i in cursor:
    # print(i)
    forging_ERP_item_id.append(i)
print("[#]Forging ERP Item ID:")
# [2000:2003] : Range of List needed by index
print("     ", forging_ERP_item_id[200:203])
# endregion  <<<<==========================[Forging ERP Item ID]===========================>>>>

# region  <<<<============================[Forging Description]============================>>>>
# Table Name : SpexForge
# Column Name : Description
cursor.execute('SELECT Description FROM SpexForge')
forging_description = []
for i in cursor:
    # print(i)
    forging_description.append(i)
print("[#]Forging Description:")
# [2000:2003] : Range of List needed by index
print("     ", forging_description[200:203])
# endregion  <<<<==========================[Forging Description]===========================>>>>

# region  <<<<============================[Forging Status]============================>>>>
# Table Name : SpexForge
# Column Name : Status
cursor.execute('SELECT Status FROM SpexForge')
forging_status = []
for i in cursor:
    # print(i)
    forging_status.append(i)
print("[#]Forging Status:")
# [2000:2003] : Range of List needed by index
print("     ", forging_status[200:203])
# endregion  <<<<==========================[Forging Status]===========================>>>>

# region  <<<<============================[Forging Bar Diameter]============================>>>>
# Table Name : SpexForge
# Column Name : BarDiameter
cursor.execute('SELECT BarDiameter FROM SpexForge')
forging_bar_diameter = []
for i in cursor:
    # print(i)
    forging_bar_diameter.append(i)
print("[#]Forging BarDiameter:")
# [2000:2003] : Range of List needed by index
print("     ", forging_bar_diameter[200:203])
# endregion  <<<<==========================[Forging Bar Diameter]===========================>>>>

# region  <<<<============================[Slug Length]============================>>>>
# Table Name : SpexForge
# Column Name : SlugLength
cursor.execute('SELECT SlugLength FROM SpexForge')
slug_length = []
for i in cursor:
    # print(i)
    slug_length.append(i)
print("[#]Slug Length:")
# [2000:2003] : Range of List needed by index
print("     ", slug_length[200:203])
# endregion  <<<<==========================[Slug Length]===========================>>>>

# region  <<<<============================[Forging Rough Dome Status]============================>>>>
# Table Name : SpexForge
# Column Name : RoughDome
cursor.execute('SELECT RoughDome FROM SpexForge')
forging_rough_dome = []
for i in cursor:
    # print(i)
    forging_rough_dome.append(i)
print("[#]Forging Rough Dome Status:")
# [2000:2003] : Range of List needed by index
print("     ", forging_rough_dome[200:203])
# endregion  <<<<==========================[Forging Rough Dome Status]===========================>>>>

# region  <<<<============================[Forging OD]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeOD
cursor.execute('SELECT ForgeOD FROM SpexForge')
forging_od = []
for i in cursor:
    # print(i)
    forging_od.append(i)
print("[#]Forging OD:")
# [2000:2003] : Range of List needed by index
print("     ", forging_od[200:203])
# endregion  <<<<==========================[Forging OD]===========================>>>>

# region  <<<<============================[Forging Rough OD Status]============================>>>>
# Table Name : SpexForge
# Column Name : RoughOD
cursor.execute('SELECT RoughOD FROM SpexForge')
forging_rough_od_status = []
for i in cursor:
    # print(i)
    forging_rough_od_status.append(i)
print("[#]Forging Rough OD Status:")
# [2000:2003] : Range of List needed by index
print("     ", forging_rough_od_status[200:203])
# endregion  <<<<==========================[Forging Rough OD Status]===========================>>>>

# region  <<<<============================[Forging Inside Diameter]============================>>>>
# Table Name : SpexForge
# Column Name : ForgeID
cursor.execute('SELECT ForgeID FROM SpexForge')
forging_inside_diameter = []
for i in cursor:
    # print(i)
    forging_inside_diameter.append(i)
print("[#]Forging Inside Diameter:")
# [2000:2003] : Range of List needed by index
print("     ", forging_inside_diameter[200:203])
# endregion  <<<<==========================[Forging Inside Diameter]===========================>>>>

# region  <<<<============================[Forging Skirt Thickness At 0]============================>>>>
# Table Name : SpexForge
# Column Name : SkirtThkAt0
cursor.execute('SELECT SkirtThkAt0 FROM SpexForge')
forging_skirt_thickness_at_0 = []
for i in cursor:
    # print(i)
    forging_skirt_thickness_at_0.append(i)
print("[#]Forging Skirt Thickness At 0:")
# [2000:2003] : Range of List needed by index
print("     ", forging_skirt_thickness_at_0[200:203])
# endregion  <<<<==========================[Forging Skirt Thickness At 0]===========================>>>>

# region  <<<<============================[Forging Skirt Thickness At 180]============================>>>>
# Table Name : SpexForge
# Column Name : SkirtThkAt180
cursor.execute('SELECT SkirtThkAt180 FROM SpexForge')
forging_skirt_thickness_at_180 = []
for i in cursor:
    # print(i)
    forging_skirt_thickness_at_180.append(i)
print("[#]Forging Skirt Thickness At 180:")
# [2000:2003] : Range of List needed by index
print("     ", forging_skirt_thickness_at_180[200:203])
# endregion  <<<<==========================[Forging Skirt Thickness At 180]===========================>>>>

# region  <<<<============================[Forging Tower Length]============================>>>>
# Table Name : SpexForge
# Column Name : TowerLength
cursor.execute('SELECT TowerLength FROM SpexForge')
forging_tower_length = []
for i in cursor:
    # print(i)
    forging_tower_length.append(i)
print("[#]Forging Tower Length:")
# [2000:2003] : Range of List needed by index
print("     ", forging_tower_length[200:203])
# endregion  <<<<==========================[Forging Tower Length]===========================>>>>

# region  <<<<============================[Forging Boss Width]============================>>>>
# Table Name : SpexForge
# Column Name : BossWidth
cursor.execute('SELECT BossWidth FROM SpexForge')
forging_boss_width = []
for i in cursor:
    # print(i)
    forging_boss_width.append(i)
print("[#]Forging Boss Width:")
# [2000:2003] : Range of List needed by index
print("     ", forging_boss_width[200:203])
# endregion  <<<<==========================[Forging Boss Width]===========================>>>>

# region  <<<<============================[Forging TBarWidth]============================>>>>
# Table Name : SpexForge
# Column Name : TBarWidth
cursor.execute('SELECT TBarWidth FROM SpexForge')
forging_TBar_width = []
for i in cursor:
    # print(i)
    forging_TBar_width.append(i)
print("[#]Forging TBar Width:")
# [2000:2003] : Range of List needed by index
print("     ", forging_TBar_width[200:203])
# endregion  <<<<==========================[Forging TBarWidth]===========================>>>>

# region  <<<<============================[Forging Boss Offset Direction]============================>>>>
# Table Name : SpexForge
# Column Name : BossOffsetDir
cursor.execute('SELECT BossOffsetDir FROM SpexForge')
forging_boss_offset_direction = []
for i in cursor:
    # print(i)
    forging_boss_offset_direction.append(i)
print("[#]Forging Boss Offset Direction:")
# [2000:2003] : Range of List needed by index
print("     ", forging_boss_offset_direction[200:203])
# endregion  <<<<==========================[Forging Boss Offset Direction]===========================>>>>

# region  <<<<============================[Forging Boss Offset Amount]============================>>>>
# Table Name : SpexForge
# Column Name : BossOffsetAmt
cursor.execute('SELECT BossOffsetAmt FROM SpexForge')
forging_boss_offset_amount = []
for i in cursor:
    # print(i)
    forging_boss_offset_amount.append(i)
print("[#]Forging Boss Offset Amount:")
# [2000:2003] : Range of List needed by index
print("     ", forging_boss_offset_amount[200:203])
# endregion  <<<<==========================[Forging Boss Offset Amount]===========================>>>>

# region  <<<<============================[Forging Inside Width At 0]============================>>>>
# Table Name : SpexForge
# Column Name : InsdForgeWidthAt0
cursor.execute('SELECT InsdForgeWidthAt0 FROM SpexForge')
forging_inside_width_at_0 = []
for i in cursor:
    # print(i)
    forging_inside_width_at_0.append(i)
print("[#]Forging Inside Width At 0:")
# [2000:2003] : Range of List needed by index
print("     ", forging_inside_width_at_0[200:203])
# endregion  <<<<==========================[Forging Inside Width At 0]===========================>>>>

# region  <<<<============================[Forging Inside Width At 180]============================>>>>
# Table Name : SpexForge
# Column Name : InsdForgeWidthAt180
cursor.execute('SELECT InsdForgeWidthAt180 FROM SpexForge')
forging_inside_width_at_180 = []
for i in cursor:
    # print(i)
    forging_inside_width_at_180.append(i)
print("[#]Forging Inside Width At 180:")
# [2000:2003] : Range of List needed by index
print("     ", forging_inside_width_at_180[200:203])
# endregion  <<<<==========================[Forging Inside Width At 180]===========================>>>>

# region  <<<<============================[Forged Dome Rise]============================>>>>
# Table Name : SpexForge
# Column Name : ForgedDomeRise
cursor.execute('SELECT ForgedDomeRise FROM SpexForge')
forged_dome_rise = []
for i in cursor:
    # print(i)
    forged_dome_rise.append(i)
print("[#]Forged Dome Rise:")
# [2000:2003] : Range of List needed by index
print("     ", forged_dome_rise[200:203])
# endregion  <<<<==========================[Forged Dome Rise]===========================>>>>

# region  <<<<============================[Forge Pilot Bore Depth]============================>>>>
# Table Name : SpexForge
# Column Name : PilotBorDpth
cursor.execute('SELECT PilotBorDpth FROM SpexForge')
forge_pilot_bore_depth = []
for i in cursor:
    # print(i)
    forge_pilot_bore_depth.append(i)
print("[#]Forge Pilot Bore Depth:")
# [2000:2003] : Range of List needed by index
print("     ", forge_pilot_bore_depth[200:203])
# endregion  <<<<==========================[Forge Pilot Bore Depth]===========================>>>>

# region  <<<<============================[Forge Pilot Bore Diameter]============================>>>>
# Table Name : SpexForge
# Column Name : PilotBorDia
cursor.execute('SELECT PilotBorDia FROM SpexForge')
forge_pilot_bore_diameter = []
for i in cursor:
    # print(i)
    forge_pilot_bore_diameter.append(i)
print("[#]Forge Pilot Bore Diameter:")
# [2000:2003] : Range of List needed by index
print("     ", forge_pilot_bore_diameter[200:203])
# endregion  <<<<==========================[Forge Pilot Bore Diameter]===========================>>>>

# region  <<<<============================[Forge Hollow Dome Rise]============================>>>>
# Table Name : SpexForge
# Column Name : HDRise
cursor.execute('SELECT HDRise FROM SpexForge')
forging_hollow_dome_rise = []
for i in cursor:
    # print(i)
    forging_hollow_dome_rise.append(i)
print("[#]Forge Hollow Dome Rise:")
# [2000:2003] : Range of List needed by index
print("     ", forging_hollow_dome_rise[200:203])
# endregion  <<<<==========================[Forge Hollow Dome Rise]===========================>>>>

# region  <<<<============================[Forge Ring Belt Height]============================>>>>
# Table Name : SpexForge
# Column Name : RingBeltHt
cursor.execute('SELECT RingBeltHt FROM SpexForge')
forge_ring_belt_height = []
for i in cursor:
    # print(i)
    forge_ring_belt_height.append(i)
print("[#]Forge Ring Belt Height:")
# [2000:2003] : Range of List needed by index
print("     ", forge_ring_belt_height[200:203])
# endregion  <<<<==========================[Forge Ring Belt Height]===========================>>>>

# region  <<<<============================[Forge Outside Width At 0]============================>>>>
# Table Name : SpexForge
# Column Name : OutsdForgeWidthAt0
cursor.execute('SELECT OutsdForgeWidthAt0 FROM SpexForge')
forge_outside_width_at_0 = []
for i in cursor:
    # print(i)
    forge_outside_width_at_0.append(i)
print("[#]Forge Outside Width At 0:")
# [2000:2003] : Range of List needed by index
print("     ", forge_outside_width_at_0[200:203])
# endregion  <<<<==========================[Forge Outside Width At 0]===========================>>>>

# region  <<<<============================[Forge Outside Width At 180]============================>>>>
# Table Name : SpexForge
# Column Name : OutsdForgeWidthAt180
cursor.execute('SELECT OutsdForgeWidthAt180 FROM SpexForge')
forge_outside_width_at_180 = []
for i in cursor:
    # print(i)
    forge_outside_width_at_180.append(i)
print("[#]Forge Outside Width At 180:")
# [2000:2003] : Range of List needed by index
print("     ", forge_outside_width_at_180[200:203])
# endregion  <<<<==========================[Forge Outside Width At 180]===========================>>>>

# region  <<<<============================[Forge Min Bore Size]============================>>>>
# Table Name : SpexForge
# Column Name : MinBoreSize
cursor.execute('SELECT MinBoreSize FROM SpexForge')
forge_min_bore_size = []
for i in cursor:
    # print(i)
    forge_min_bore_size.append(i)
print("[#]Forge Min Bore Size:")
# [2000:2003] : Range of List needed by index
print("     ", forge_min_bore_size[200:203])
# endregion  <<<<==========================[Forge Min Bore Size]===========================>>>>

# region  <<<<============================[Forge Max Bore Size]============================>>>>
# Table Name : SpexForge
# Column Name : MaxBoreSize
cursor.execute('SELECT MaxBoreSize FROM SpexForge')
forge_max_bore_size = []
for i in cursor:
    # print(i)
    forge_max_bore_size.append(i)
print("[#]Forge Max Bore Size:")
# [2000:2003] : Range of List needed by index
print("     ", forge_max_bore_size[200:203])
# endregion  <<<<==========================[Forge Max Bore Size]===========================>>>>

# region  <<<<============================[Forge Parts Per Bar]============================>>>>
# Table Name : SpexForge
# Column Name : PartsPerBar
cursor.execute('SELECT PartsPerBar FROM SpexForge')
forge_parts_per_bar = []
for i in cursor:
    # print(i)
    forge_parts_per_bar.append(i)
print("[#]Forge Parts Per Bar:")
# [2000:2003] : Range of List needed by index
print("     ", forge_parts_per_bar[200:203])
# endregion  <<<<==========================[Forge Parts Per Bar]===========================>>>>

#          <<<<<<<<------------------------------------------------------------------------>>>>>>>>>>
#      <<<<<<<<--------------------------------------------------------------------------------->>>>>>>>>>
# <<<<<<<<------------------------------------------------------------------------------------------->>>>>>>>>>
# endregion  <<<<===========================[Forging Information from DataBase]===========================>>>>

print()
# region  <<<<===================[Template to Get Specific Value For Specific Forging]===================>>>>
# Change Column Names according to value you looking for
# Change 'PistonID' to check different job number

# **If need to get ForgingDimension by giving the forging number direct.
# forging_number_test = 'FJE89M-HEX'
# cursor.execute('SELECT ForgeOD FROM SpexForge WHERE ForgeItemID = ?', forging_number_test)

# **If need to get ForgingDimension for the forging number that use for the job number above.
cursor.execute('SELECT BossInsdSpace FROM SpexForge WHERE ForgeItemID = ?', forging_number)
specific_value_for_specific_forging = []
for i in cursor.fetchone():
    # print(i)
    specific_value_for_specific_forging = i
# **If need to print ForgingDimension by giving the forging number direct.
# print("[#]specific_value_for_specific_forging " + forging_number_test + ":")
# **If need to print ForgingDimension for the forging number that use for the job number above.
print("[#]specific_value_for_specific_forging " + forging_number + ":")
# [25000:25003] : Range of List needed by index
print("     ", specific_value_for_specific_forging)
# endregion  <<<<=================[Template to Get Specific Value For Specific Forging]=================>>>>
