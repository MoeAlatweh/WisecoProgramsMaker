import glob

search_path = 'H:\CNCProgs\HOREBORE\Probe Programs'
filename = 'FJE58MC'

RESULT_OF_PROBE_PROGRAM_SEARCH = []
# for file in glob.glob("H:\CNCProgs\*\*WD-12756*"):
for file in glob.glob(search_path+'*\*'+filename + '*'):
    # print(file)
    RESULT_OF_PROBE_PROGRAM_SEARCH.append(file)
    # print(RESULT_OF_PROBE_PROGRAM_SEARCH)
    # print(len(RESULT_OF_PROBE_PROGRAM_SEARCH))
# DECLARE AN EMPTY LIST TO ADD PROBE PROGRAM LINES ONE BY ONE.
PROBE_PROGRAMS_LINES = []
if ((len(RESULT_OF_PROBE_PROGRAM_SEARCH) == 1) and (filename != '')):
    print(RESULT_OF_PROBE_PROGRAM_SEARCH[0])
    with open(RESULT_OF_PROBE_PROGRAM_SEARCH[0], 'rt') as CurrentProgram:
        for line in CurrentProgram:  # For each line in the file,
            PROBE_PROGRAMS_LINES.append(line.rstrip('\n'))  # strip newline and add to list.
        # print(PROBE_PROGRAMS_LINES)
        print("LINE OF PROBE PROGRAM THAT NEED TO ADD TO HORIZONTAL PROGRAM: "+PROBE_PROGRAMS_LINES[0])
elif ((len(RESULT_OF_PROBE_PROGRAM_SEARCH) == 0) and (filename != '')):
    print("NO PROBE PROGRAM FOUND, MAKE ONE AND CLICK SUBMIT BUTTON AGAIN ")
elif ((len(RESULT_OF_PROBE_PROGRAM_SEARCH) > 1) and (filename != '')):
    # NEED TO WORK ON THAT
    print("MORE RESULTS FOUND , CHOOSE THE RIGHT PROBE PROGRAM")
    print(RESULT_OF_PROBE_PROGRAM_SEARCH)
    print(len(RESULT_OF_PROBE_PROGRAM_SEARCH))
else:
    print("SOMETHING UNEXPECTED HAPPEN, SEE PROGRAMMING ")
