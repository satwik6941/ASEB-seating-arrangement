import pandas as pd

df = pd.read_excel("Subjects (2).xlsx", sheet_name="names")

a = df['Roll No']

CSE_year_1 = []
CSE_year_2 = []
CSE_year_3 = []
CSE_year_4 = []
AIE_year_1 = []
AIE_year_2 = []
AIE_year_3 = []
AIE_year_4 = []
AID_year_1 = []
AID_year_2 = []
AID_year_3 = []
AID_year_4 = []
ECE_year_1 = []
ECE_year_2 = []
ECE_year_3 = []
ECE_year_4 = []
EAC_year_1 = []
EAC_year_2 = []
EAC_year_3 = []
EAC_year_4 = []
ELC_year_1 = []
ELC_year_2 = []
ELC_year_3 = []
ELC_year_4 = []
EEE_year_1 = []
EEE_year_2 = [] 
EEE_year_3 = []
EEE_year_4 = []
MEE_year_1 = []
MEE_year_2 = []
MEE_year_3 = []
MEE_year_4 = []
RAE_year_1 = [] 
RAE_year_2 = [] 
RAE_year_3 = [] 
RAE_year_4 = []

for reg_no in df["Roll No"]:
    if reg_no[8:13] == 'CSE24':
        CSE_year_1.append(reg_no)
    elif reg_no[8:13] == 'CSE23':
        CSE_year_2.append(reg_no)
    elif reg_no[8:13] == 'CSE22':
        CSE_year_3.append(reg_no)
    elif reg_no[8:13] == 'CSE21':
        CSE_year_4.append(reg_no)
    elif reg_no[8:13] == 'AIE24':
        AIE_year_1.append(reg_no)
    elif reg_no[8:13] == 'AIE23':
        AIE_year_2.append(reg_no)
    elif reg_no[8:13] == 'AIE22':
        AIE_year_3.append(reg_no)
    elif reg_no[8:13] == 'AIE21':
        AIE_year_4.append(reg_no)
    elif reg_no[8:13] == 'AID24':
        AID_year_1.append(reg_no)
    elif reg_no[8:13] == 'AID23':
        AID_year_2.append(reg_no)
    elif reg_no[8:13] == 'AID22':
        AID_year_3.append(reg_no)
    elif reg_no[8:13] == 'AID21':
        AID_year_4.append(reg_no)
    elif reg_no[8:13] == 'ECE24':
        ECE_year_1.append(reg_no)
    elif reg_no[8:13] == 'ECE23':
        ECE_year_2.append(reg_no)
    elif reg_no[8:13] == 'ECE22':
        ECE_year_3.append(reg_no)
    elif reg_no[8:13] == 'ECE21':
        ECE_year_4.append(reg_no)
    elif reg_no[8:13] == 'EAC24':
        EAC_year_1.append(reg_no)
    elif reg_no[8:13] == 'EAC23':
        EAC_year_2.append(reg_no)
    elif reg_no[8:13] == 'EAC22':
        EAC_year_3.append(reg_no)
    elif reg_no[8:13] == 'EAC21':
        EAC_year_4.append(reg_no)
    elif reg_no[8:13] == 'ELC24':
        ELC_year_1.append(reg_no)
    elif reg_no[8:13] == 'ELC23':
        ELC_year_2.append(reg_no)
    elif reg_no[8:13] == 'ELC22':
        ELC_year_3.append(reg_no)
    elif reg_no[8:13] == 'ELC21':
        ELC_year_4.append(reg_no)
    elif reg_no[8:13] == 'EEE24':
        EEE_year_1.append(reg_no)
    elif reg_no[8:13] == 'EEE23':
        EEE_year_2.append(reg_no)
    elif reg_no[8:13] == 'EEE22':
        EEE_year_3.append(reg_no)
    elif reg_no[8:13] == 'EEE21':
        EEE_year_4.append(reg_no)
    elif reg_no[8:13] == 'MEE24':
        MEE_year_1.append(reg_no)
    elif reg_no[8:13] == 'MEE23':
        MEE_year_2.append(reg_no)
    elif reg_no[8:13] == 'MEE22':
        MEE_year_3.append(reg_no)
    elif reg_no[8:13] == 'MEE21':
        MEE_year_4.append(reg_no)
    elif reg_no[8:13] == 'RAE24':
        RAE_year_1.append(reg_no)
    elif reg_no[8:13] == 'RAE23':
        RAE_year_2.append(reg_no)
    elif reg_no[8:13] == 'RAE22':
        RAE_year_3.append(reg_no)
    elif reg_no[8:13] == 'RAE21':
        RAE_year_4.append(reg_no)

print("CSE Year 1:", len(CSE_year_1))
print("CSE Year 2:", len(CSE_year_2))
print("CSE Year 3:", len(CSE_year_3))
print("CSE Year 4:", len(CSE_year_4))
print("AIE Year 1:", len(AIE_year_1))
print("AIE Year 2:", len(AIE_year_2))
print("AIE Year 3:", len(AIE_year_3))
print("AIE Year 4:", len(AIE_year_4))
print("AID Year 1:", len(AID_year_1))
print("AID Year 2:", len(AID_year_2))
print("AID Year 3:", len(AID_year_3))
print("AID Year 4:", len(AID_year_4))
print("ECE Year 1:", len(ECE_year_1))
print("ECE Year 2:", len(ECE_year_2))
print("ECE Year 3:", len(ECE_year_3))
print("ECE Year 4:", len(ECE_year_4))
print("EAC Year 1:", len(EAC_year_1))
print("EAC Year 2:", len(EAC_year_2))
print("EAC Year 3:", len(EAC_year_3))
print("EAC Year 4:", len(EAC_year_4))
print("ELC Year 1:", len(ELC_year_1))
print("ELC Year 2:", len(ELC_year_2))
print("ELC Year 3:", len(ELC_year_3))
print("ELC Year 4:", len(ELC_year_4))
print("EEE Year 1:", len(EEE_year_1))
print("EEE Year 2:", len(EEE_year_2))
print("EEE Year 3:", len(EEE_year_3))
print("EEE Year 4:", len(EEE_year_4))
print("MEE Year 1:", len(MEE_year_1))
print("MEE Year 2:", len(MEE_year_2))
print("MEE Year 3:", len(MEE_year_3))
print("MEE Year 4:", len(MEE_year_4))
print("RAE Year 1:", len(RAE_year_1))
print("RAE Year 2:", len(RAE_year_2))
print("RAE Year 3:", len(RAE_year_3))
print("RAE Year 4:", len(RAE_year_4))
