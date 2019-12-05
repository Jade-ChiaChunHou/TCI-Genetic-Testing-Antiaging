### Part 0: Import

import os
import numpy as np
import pandas as pd
from statistics import mean

### Part 1: Read Files
#Raw data:
#(1) cycle 1 & cycle 35
#(2) genotype

################################
# Read files: cycle 1 & cycle 35
################################

AG_num = str(input("Please the AG sample number(ex: AG190604_A11_B12): "))
sample_num = str(input("Please enter the sample number(ex: A, B, C, D): "))
#AG_num = 'AG190604_A11_B12'

cycle1_dic = '../sample_data/' + AG_num +'-cycle 1.xls'
cycle35_dic = '../sample_data/' + AG_num + '-cycle 35.xls'


cycle1 = pd.read_excel(open(cycle1_dic, 'rb'),
              sheet_name='Multicomponent Data',
              header=None,
              names = ['Well', 'Cycle', 'ROX', 'VIC', 'FAM'])

cycle1 = pd.DataFrame(cycle1)
cycle1 = cycle1.loc[47:]
cycle1.head(50)

cycle35 = pd.read_excel(open(cycle35_dic, 'rb'),
              sheet_name='Multicomponent Data',
              header=None,
              names = ['Well', 'Cycle', 'ROX', 'VIC', 'FAM'])


cycle35 = pd.DataFrame(cycle35)
cycle35 = cycle35.loc[47:]
cycle35.head(50)


########################
# Read file: AG genotype
########################

ag_genotype = pd.read_csv('../raw/AG_genotype.csv')

# genotype 1
gene1 = ag_genotype.iloc[:,7]

# genotype 2
gene2 = ag_genotype.iloc[:,8]

# genotype 3
gene3 = ag_genotype.iloc[:,9]

#standard 1
stand1 = ag_genotype.iloc[:,13]

#standard 2
stand2 = ag_genotype.iloc[:,14]

# gene risk 1
risk1 = ag_genotype.iloc[:,10]

# gene risk 2
risk2 = ag_genotype.iloc[:,11]

# gene risk 3
risk3 = ag_genotype.iloc[:,12]

ag_genotype.head(87)


###################
# Read file: member
###################


member = pd.read_excel(open('../raw/來案暨出貨明細表-基因檢測.xlsx', 'rb'),
              sheet_name='來案暨出貨明細表-基因檢測(201904起')


# get sample number for membership

if sample_num.upper() == "A":

    index = AG_num[8:].find("A") + 8
    num = AG_num[(index + 1): (index + 3)]

    AG_sample = AG_num[0:8] + '-' + num


elif sample_num.upper() == "B":

    index = AG_num[8:].find("B") + 8
    num = AG_num[(index + 1): (index + 3)]

    AG_sample = AG_num[0:8] + '-' + num


elif sample_num.upper() == "C":

    index = AG_num[8:].find("C") + 8
    num = AG_num[(index + 1): (index + 3)]

    AG_sample = AG_num[0:8] + '-' + num


elif sample_num.upper() == "D":

    index = AG_num[8:].find("D") + 8
    num = AG_num[(index + 1): (index + 3)]

    AG_sample = AG_num[0:8] + '-' + num


'''
member[(member[['檢體編號']] == AG_sample).all(1)]

# member number
member_number = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,10]

# report date
date = str(member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,9])[0:4] + '/' + str(member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,9])[5:7] + '/' +str(member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,9])[8:10]

# gender
gender = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,13]

# paper or electronic
report_type = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,16]

# language
language = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,17]


# clinic

if member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] == 0:
    clinic = "健樂士診所"

#elif math.isnan(member[(member[['檢體編號']] == G2_num).all(1)].iloc[0,6]):
#    clinic = "健樂士診所"

elif member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] == '':
    clinic = "健樂士診所"

elif member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] != '':
    clinic = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6]



# clinic SC

if member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] == 0:
    clinic_SC = "健乐士诊所"

#elif math.isnan(member[(member[['檢體編號']] == G2_num).all(1)].iloc[0,6]):
#    clinic_SC = "健乐士诊所"

elif member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] == '':
    clinic_SC = "健乐士诊所"

elif member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6] != '':
    clinic_SC = member[(member[['檢體編號']] == AG_sample).all(1)].iloc[0,6]


membership = {'data':['sample ID', 'clinic', 'member ID', 'date', 'gender', 'report_type', 'language'], 'info':[AG_sample, clinic, member_number, date, gender, report_type, language]}

membership_SC = {'data':['sample ID', 'clinic', 'member ID', 'date', 'gender', 'report_type', 'language'], 'info':[AG_sample, clinic_SC, member_number, date, gender, report_type, language]}

# Create DataFrame
membership = pd.DataFrame(membership)
membership

# write csv
outname = AG_sample + '_membership.csv'

outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
if not os.path.exists(outdir):
    os.makedirs(outdir)

fullname = os.path.join(outdir, outname)

membership.to_csv(fullname , encoding='utf-8')

print("(1) " +AG_sample + ' finish output ' + fullname + '\n')


# membership_SC

membership_SC = pd.DataFrame(membership_SC)
membership_SC

# write csv
outname = AG_sample + '_membership_SC.csv'

outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
if not os.path.exists(outdir):
    os.makedirs(outdir)

fullname = os.path.join(outdir, outname)

membership_SC.to_csv(fullname , encoding='utf-8')


print("(2) " + AG_sample + ' finish output ' + fullname + '\n')

print("---------------------------")
print('sample ID: ', AG_sample)
print('clinic: ', clinic)
print('member ID: ', member_number)
print('report date: ', date)
print('gender: ', gender)
print('report_type: ', report_type)
print('language: ', language)
print("---------------------------")
print("")

membership
'''

### Part 2: Calculate the vic & fam

#################################
# calculate delta_vic & delta_fam

delta_vic = list(cycle35['VIC'] - cycle1['VIC'])
#print(delta_vic)

delta_fam = list(cycle35['FAM'] - cycle1['FAM'])
#print(delta_fam)

delta = pd.DataFrame(list(zip(delta_vic, delta_fam)),
                    columns=['delta_vic','delta_fam'])



delta.head()

print('(3) ' + AG_num + ' get vic and fam value.\n')



#############################
# get the signal of delta_vic

vic_signal = []

for cell in delta_vic:
    if cell >= 500:
        vic_signal.append(1)

    else:
        vic_signal.append(0)

#print(vic_signal)


print('(4) ' + AG_num + ' finish delta_vic calculation.\n')


#############################
# get the signal of delta_fam

fam_signal = []

for cell in delta_fam:
    if cell >= 500:
        fam_signal.append(1)

    else:
        fam_signal.append(0)

#print(fam_signal)


print('(5) ' + AG_num + ' finish delta_fam calculation.\n')


########################################
# get the error of delta_vic & delta_fam

signal = []

for i in range(0,len(vic_signal)):
    signal.append(vic_signal[i] + fam_signal[i])

#print(signal)


error = []

for cell in signal:
    if cell == 2:
        error.append(0)
    elif cell != 2:
        error.append(1)



######################################
# get the value: delta_vic / delta_fam

value = []

for i in range(len(delta_vic )):

    if vic_signal[i] == 1 and fam_signal[i] == 1:

        value.append(delta_vic[i] / delta_fam[i])

    elif vic_signal[i] != 1 and fam_signal[i] == 1:

        value.append(1 / delta_fam[i])

    elif vic_signal[i] == 1 and fam_signal[i] != 1:

        value.append(delta_vic[i] / 1)

    else:
        value.append(1)

#print(value)

print('(6) ' + AG_num + ' finish delta vic and fam ratio calculate.\n')



value_vic_fam = pd.DataFrame(list(zip(delta_vic, delta_fam, vic_signal, fam_signal, signal, error, value)),
                    columns=['delta_vic','delta_fam', 'vic_signal', 'fam_signal', 'sum_signal', 'error', 'value'])


### Part 3: Calculate the genotype: A-D


if sample_num.upper() == "A":

    # get value (delta vic / delta fam)

    a_value = value_vic_fam.iloc[[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
                        24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
                        48, 49,
                        35,
                        50, 51, 52, 53, 54, 55, 56,
                        51, 50,
                        57, 58, 59,
                        72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83,
                        96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107,
                        120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131,
                        144, 145, 146, 147, 148, 149, 150, 151, 153, 152, 154, 155,
                        168, 169, 170, 171, 172, 173, 174, 175, 176, 177], [6]]

    a_ratio = []

    for row in range(len(a_value)):

        a_ratio.append(a_value.iloc[row,0])

    print("(7) " + AG_num + ' finish calculate the "a_value".' + '\n')

    #a_ratio


    # Get genotype

    a_genotype = []

    for row in range(len(a_value)):

        if stand1[row] > stand2[row]:

            if float(a_value.iloc[row,]) > stand1[row]:

                a_genotype.append(gene1[row])

            elif float(a_value.iloc[row,]) < stand2[row]:

                a_genotype.append(gene3[row])

            else:

                a_genotype.append(gene2[row])

        elif stand1[row] < stand2[row]:

            if float(a_value.iloc[row,]) > stand2[row]:

                a_genotype.append(gene3[row])

            elif float(a_value.iloc[row,]) < stand1[row]:

                a_genotype.append(gene1[row])

            else:

                a_genotype.append(gene2[row])


    print("(8) " + AG_num + ' finish calculate "a_genotype".' +'\n')

    #a_genotype



    # A score

    a_score = []

    for row in range(len(a_genotype)):

        if a_genotype[row] == gene1[row]:

            a_score.append(40*risk1[row] + 20)

        elif a_genotype[row] == gene2[row]:

            a_score.append(20*risk2[row] + 60)

        elif a_genotype[row] == gene3[row]:

            a_score.append(20*risk3[row] + 80)

    print("(9) " + AG_num + ' finish calculate "a_score". ' +'\n')


    # A result

    #############
    # (1) 細胞衰老

    # 1.細胞循環

    g1_1 = mean(a_score[0:5])

    # 2.DNA損傷檢查點

    g1_2 = mean(a_score[5:10])

    # 3.端粒酶活性

    g1_3 = mean(a_score[10:12])

    # 4.DNA修復

    g1_4 = mean(a_score[12:20])

    # 5.粒線體活性

    g1_5 = mean(a_score[20:26])

    # 6.對抗自由基

    g1_6 = mean(a_score[26:32])


    #############
    # (2) 毒物代謝

    # 1.肝臟

    g2_1 = mean(a_score[32:37])

    # 2.肺臟

    g2_2 = mean(a_score[37:40])

    # 3.腸道

    g2_3 = mean(a_score[40:43])

    # 4.腎臟

    g2_4 = mean(a_score[43:47])


    #############
    # (3) 營養代謝

    # 1.血管健康

    g3_1 = mean(a_score[47:51])

    # 2.蛋白質代謝

    g3_2 = mean(a_score[51:57])

    # 3.醣類代謝

    g3_3 = mean(a_score[57:65])

    # 4.脂質代謝

    g3_4 = mean(a_score[65:72])


    #############
    # (4) 免疫調節

    # 1.抗原呈現

    g4_1 = mean(a_score[72:77])

    # 2.免疫訓練

    g4_2 = mean(a_score[77:88])

    # 3.發炎因子

    g4_3 = mean(a_score[88:97])



    #####################
    # Score of 4 category

    # (1) 細胞衰老

    g1 = mean(a_score[0:32])

    # (2) 毒物代謝

    g2 = mean(a_score[32:47])

    # (3) 營養代謝

    g3 = mean(a_score[47:72])

    # (4) 免疫調節

    g4 = mean(a_score[72:97])


    #############################################
    # Pick the bad subcategory from each category


    sub = []

    # (1) 細胞衰老

    if min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_1:

        sub.append("細胞循環")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_2:

        sub.append("DNA損傷檢查點")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_3:

        sub.append("端粒酶活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_4:

        sub.append("DNA修復")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_5:

        sub.append("粒線體活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_6:

        sub.append("對抗自由基")


    # (2) 毒物代謝

    if min(g2_1, g2_2, g2_3, g2_4) == g2_1:

        sub.append("肝臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_2:

        sub.append("肺臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_3:

        sub.append("腸道")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_4:

        sub.append("腎臟")


    # (3) 營養代謝

    if min(g3_1, g3_2, g3_3, g3_4) == g3_1:

        sub.append("血管健康")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_2:

        sub.append("蛋白質代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_3:

        sub.append("醣類代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_4:

        sub.append("脂質代謝")


    # (4) 免疫調節

    if min(g4_1, g4_2, g4_3) == g4_1:

        sub.append("抗原呈現")

    elif min(g4_1, g4_2, g4_3) == g4_2:

        sub.append("免疫訓練")

    elif min(g4_1, g4_2, g4_3) == g4_3:

        sub.append("發炎因子")


    print("(10) " + AG_num + ' finish calculate "a_result". ' +'\n')

    sub

    #####################
    # output a_result.csv

    ag_genotype['a_value'] = a_ratio
    ag_genotype['a_genotype'] = a_genotype
    ag_genotype['a_score'] = a_score

    ag_genotype


    outname = AG_num + '_a_result.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    ag_genotype.to_csv(fullname)

    print("(11) " + AG_num + ' finish output ' + fullname + '\n')


    #######################
    # output a_category.csv

    category = ["細胞衰老", "毒物代謝", "營養代謝", "免疫調節"]

    a_mean = [format(g1, '.0f'), format(g2, '.0f'), format(g3, '.0f'), format(g4, '.0f')]

    a_category = pd.DataFrame(list(zip(category, a_mean, sub)),
                  columns=['category','mean', 'high risk subcategory'])


    outname = AG_num + '_a_category.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    a_category.to_csv(fullname)

    print("(12) " + AG_num + ' finish output ' + fullname + '\n')



elif sample_num.upper() == "B":

    # get value (delta vic / delta fam)

    b_value = value_vic_fam.iloc[[12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
                        36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47,
                        60, 61,
                        47,
                        62, 63, 64, 65, 66, 67, 68,
                        63, 62,
                        69, 70, 71,
                        84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95,
                        108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119,
                        132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143,
                        156, 157, 158, 159, 160, 161, 162, 163, 165, 164, 166, 167,
                        180, 181, 182, 183, 184, 185, 186, 187, 188, 189], [6]]

    b_ratio = []

    for row in range(len(b_value)):

        b_ratio.append(b_value.iloc[row,0])

    print(AG_num + ' finish calculate the "b_value".' + '\n')



    # Get genotype

    b_genotype = []

    for row in range(len(b_value)):

        if stand1[row] > stand2[row]:

            if float(b_value.iloc[row,]) > stand1[row]:

                b_genotype.append(gene1[row])

            elif float(b_value.iloc[row,]) < stand2[row]:

                b_genotype.append(gene3[row])

            else:

                b_genotype.append(gene2[row])

        elif stand1[row] < stand2[row]:

            if float(b_value.iloc[row,]) > stand2[row]:

                b_genotype.append(gene3[row])

            elif float(b_value.iloc[row,]) < stand1[row]:

                b_genotype.append(gene1[row])

            else:

                b_genotype.append(gene2[row])


    print(AG_num + ' finish calculate "b_genotype".' +'\n')

    #b_genotype



    # B score

    b_score = []

    for row in range(len(b_genotype)):

        if b_genotype[row] == gene1[row]:

            b_score.append(40*risk1[row] + 20)

        elif b_genotype[row] == gene2[row]:

            b_score.append(20*risk2[row] + 60)

        elif b_genotype[row] == gene3[row]:

            b_score.append(20*risk3[row] + 80)

    print(AG_num + ' finish calculate "b_score". ' +'\n')



    # B result

    #############
    # (1) 細胞衰老

    # 1.細胞循環

    g1_1 = mean(b_score[0:5])

    # 2.DNA損傷檢查點

    g1_2 = mean(b_score[5:10])

    # 3.端粒酶活性

    g1_3 = mean(b_score[10:12])

    # 4.DNA修復

    g1_4 = mean(b_score[12:20])

    # 5.粒線體活性

    g1_5 = mean(b_score[20:26])

    # 6.對抗自由基

    g1_6 = mean(b_score[26:32])


    #############
    # (2) 毒物代謝

    # 1.肝臟

    g2_1 = mean(b_score[32:37])

    # 2.肺臟

    g2_2 = mean(b_score[37:40])

    # 3.腸道

    g2_3 = mean(b_score[40:43])

    # 4.腎臟

    g2_4 = mean(b_score[43:47])


    #############
    # (3) 營養代謝

    # 1.血管健康

    g3_1 = mean(b_score[47:51])

    # 2.蛋白質代謝

    g3_2 = mean(b_score[51:57])

    # 3.醣類代謝

    g3_3 = mean(b_score[57:65])

    # 4.脂質代謝

    g3_4 = mean(b_score[65:72])


    #############
    # (4) 免疫調節

    # 1.抗原呈現

    g4_1 = mean(b_score[72:77])

    # 2.免疫訓練

    g4_2 = mean(b_score[77:88])

    # 3.發炎因子

    g4_3 = mean(b_score[88:97])



    #####################
    # Score of 4 category

    # (1) 細胞衰老

    g1 = mean(b_score[0:32])

    # (2) 毒物代謝

    g2 = mean(b_score[32:47])

    # (3) 營養代謝

    g3 = mean(b_score[47:72])

    # (4) 免疫調節

    g4 = mean(b_score[72:97])


    #############################################
    # Pick the bad subcategory from each category


    sub = []

    # (1) 細胞衰老

    if min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_1:

        sub.append("細胞循環")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_2:

        sub.append("DNA損傷檢查點")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_3:

        sub.append("端粒酶活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_4:

        sub.append("DNA修復")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_5:

        sub.append("粒線體活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_6:

        sub.append("對抗自由基")


    # (2) 毒物代謝

    if min(g2_1, g2_2, g2_3, g2_4) == g2_1:

        sub.append("肝臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_2:

        sub.append("肺臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_3:

        sub.append("腸道")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_4:

        sub.append("腎臟")


    # (3) 營養代謝

    if min(g3_1, g3_2, g3_3, g3_4) == g3_1:

        sub.append("血管健康")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_2:

        sub.append("蛋白質代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_3:

        sub.append("醣類代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_4:

        sub.append("脂質代謝")


    # (4) 免疫調節

    if min(g4_1, g4_2, g4_3) == g4_1:

        sub.append("抗原呈現")

    elif min(g4_1, g4_2, g4_3) == g4_2:

        sub.append("免疫訓練")

    elif min(g4_1, g4_2, g4_3) == g4_3:

        sub.append("發炎因子")


    print(AG_num + ' finish calculate "b_result". ' +'\n')

    sub

    #####################
    # output b_result.csv

    ag_genotype['b_value'] = b_ratio
    ag_genotype['b_genotype'] = b_genotype
    ag_genotype['b_score'] = b_score

    ag_genotype


    outname = AG_num + '_b_result.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    ag_genotype.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')


    #######################
    # output b_category.csv

    category = ["細胞衰老", "毒物代謝", "營養代謝", "免疫調節"]

    b_mean = [format(g1, '.0f'), format(g2, '.0f'), format(g3, '.0f'), format(g4, '.0f')]

    b_category = pd.DataFrame(list(zip(category, b_mean, sub)),
                  columns=['category','mean', 'high risk subcategory'])


    outname = AG_num + '_b_category.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    b_category.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')



elif sample_num.upper() == "C":

    # get value (delta vic / delta fam)

    c_value = value_vic_fam.iloc[[192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203,
                        216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227,
                        240, 241,
                        227,
                        242, 243, 244, 245, 246, 247, 248,
                        243, 242,
                        249, 250, 251,
                        264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275,
                        288, 289, 290, 291, 292, 293, 294, 295, 296, 297, 298, 299,
                        312, 313, 314, 315, 316, 317, 318, 319, 320, 321, 322, 323,
                        336, 337, 338, 339, 340, 341, 342, 343, 345, 344, 346, 347,
                        360, 361, 362, 363, 364, 365, 366, 367, 368, 369], [6]]

    c_ratio = []

    for row in range(len(c_value)):

        c_ratio.append(c_value.iloc[row,0])

    print(AG_num + ' finish calculate the "c_value".' + '\n')


    # Get genotype

    c_genotype = []

    for row in range(len(c_value)):

        if stand1[row] > stand2[row]:

            if float(c_value.iloc[row,]) > stand1[row]:

                c_genotype.append(gene1[row])

            elif float(c_value.iloc[row,]) < stand2[row]:

                c_genotype.append(gene3[row])

            else:

                c_genotype.append(gene2[row])

        elif stand1[row] < stand2[row]:

            if float(c_value.iloc[row,]) > stand2[row]:

                c_genotype.append(gene3[row])

            elif float(c_value.iloc[row,]) < stand1[row]:

                c_genotype.append(gene1[row])

            else:

                c_genotype.append(gene2[row])


    print(AG_num + ' finish calculate "c_genotype".' +'\n')

    #c_genotype



    # C score

    c_score = []

    for row in range(len(c_genotype)):

        if c_genotype[row] == gene1[row]:

            c_score.append(40*risk1[row] + 20)

        elif c_genotype[row] == gene2[row]:

            c_score.append(20*risk2[row] + 60)

        elif c_genotype[row] == gene3[row]:

            c_score.append(20*risk3[row] + 80)

    print(AG_num + ' finish calculate "c_score". ' +'\n')



    # C result

    #############
    # (1) 細胞衰老

    # 1.細胞循環

    g1_1 = mean(c_score[0:5])

    # 2.DNA損傷檢查點

    g1_2 = mean(c_score[5:10])

    # 3.端粒酶活性

    g1_3 = mean(c_score[10:12])

    # 4.DNA修復

    g1_4 = mean(c_score[12:20])

    # 5.粒線體活性

    g1_5 = mean(c_score[20:26])

    # 6.對抗自由基

    g1_6 = mean(c_score[26:32])


    #############
    # (2) 毒物代謝

    # 1.肝臟

    g2_1 = mean(c_score[32:37])

    # 2.肺臟

    g2_2 = mean(c_score[37:40])

    # 3.腸道

    g2_3 = mean(c_score[40:43])

    # 4.腎臟

    g2_4 = mean(c_score[43:47])


    #############
    # (3) 營養代謝

    # 1.血管健康

    g3_1 = mean(c_score[47:51])

    # 2.蛋白質代謝

    g3_2 = mean(c_score[51:57])

    # 3.醣類代謝

    g3_3 = mean(c_score[57:65])

    # 4.脂質代謝

    g3_4 = mean(c_score[65:72])


    #############
    # (4) 免疫調節

    # 1.抗原呈現

    g4_1 = mean(c_score[72:77])

    # 2.免疫訓練

    g4_2 = mean(c_score[77:88])

    # 3.發炎因子

    g4_3 = mean(c_score[88:97])



    #####################
    # Score of 4 category

    # (1) 細胞衰老

    g1 = mean(c_score[0:32])

    # (2) 毒物代謝

    g2 = mean(c_score[32:47])

    # (3) 營養代謝

    g3 = mean(c_score[47:72])

    # (4) 免疫調節

    g4 = mean(c_score[72:97])


    #############################################
    # Pick the bad subcategory from each category


    sub = []

    # (1) 細胞衰老

    if min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_1:

        sub.append("細胞循環")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_2:

        sub.append("DNA損傷檢查點")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_3:

        sub.append("端粒酶活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_4:

        sub.append("DNA修復")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_5:

        sub.append("粒線體活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_6:

        sub.append("對抗自由基")


    # (2) 毒物代謝

    if min(g2_1, g2_2, g2_3, g2_4) == g2_1:

        sub.append("肝臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_2:

        sub.append("肺臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_3:

        sub.append("腸道")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_4:

        sub.append("腎臟")


    # (3) 營養代謝

    if min(g3_1, g3_2, g3_3, g3_4) == g3_1:

        sub.append("血管健康")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_2:

        sub.append("蛋白質代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_3:

        sub.append("醣類代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_4:

        sub.append("脂質代謝")


    # (4) 免疫調節

    if min(g4_1, g4_2, g4_3) == g4_1:

        sub.append("抗原呈現")

    elif min(g4_1, g4_2, g4_3) == g4_2:

        sub.append("免疫訓練")

    elif min(g4_1, g4_2, g4_3) == g4_3:

        sub.append("發炎因子")


    print(AG_num + ' finish calculate "c_result". ' +'\n')

    sub

    #####################
    # output c_result.csv

    ag_genotype['c_value'] = c_ratio
    ag_genotype['c_genotype'] = c_genotype
    ag_genotype['c_score'] = c_score

    ag_genotype


    outname = AG_num + '_c_result.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    ag_genotype.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')


    #######################
    # output c_category.csv

    category = ["細胞衰老", "毒物代謝", "營養代謝", "免疫調節"]

    c_mean = [format(g1, '.0f'), format(g2, '.0f'), format(g3, '.0f'), format(g4, '.0f')]

    c_category = pd.DataFrame(list(zip(category, c_mean, sub)),
                  columns=['category','mean', 'high risk subcategory'])


    outname = AG_num + '_c_category.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    c_category.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')





elif sample_num.upper() == "D":

    # get value (delta vic / delta fam)

    d_value = value_vic_fam.iloc[[204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215,
                        228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239,
                        252, 253,
                        239,
                        254, 255, 256, 257, 258, 259, 260,
                        255, 254,
                        261, 262, 263,
                        276, 277, 278, 279, 280, 281, 282, 283, 284, 285, 286, 287,
                        300, 301, 302, 303, 304, 305, 306, 307, 308, 309, 310, 311,
                        324, 325, 326, 327, 328, 329, 330, 331, 332, 333, 334, 335,
                        348, 349, 350, 351, 352, 353, 354, 355, 357, 356, 358, 359,
                        372, 373, 374, 375, 376, 377, 378, 379, 380, 381], [6]]

    d_ratio = []

    for row in range(len(d_value)):

        d_ratio.append(d_value.iloc[row,0])

    print(AG_num + ' finish calculate the "d_value".' + '\n')


    # Get genotype

    d_genotype = []

    for row in range(len(d_value)):

        if stand1[row] > stand2[row]:

            if float(d_value.iloc[row,]) > stand1[row]:

                d_genotype.append(gene1[row])

            elif float(d_value.iloc[row,]) < stand2[row]:

                d_genotype.append(gene3[row])

            else:

                d_genotype.append(gene2[row])

        elif stand1[row] < stand2[row]:

            if float(d_value.iloc[row,]) > stand2[row]:

                d_genotype.append(gene3[row])

            elif float(d_value.iloc[row,]) < stand1[row]:

                d_genotype.append(gene1[row])

            else:

                d_genotype.append(gene2[row])


    print(AG_num + ' finish calculate "d_genotype".' +'\n')

    #d_genotype



    # D score

    d_score = []

    for row in range(len(d_genotype)):

        if d_genotype[row] == gene1[row]:

            d_score.append(40*risk1[row] + 20)

        elif d_genotype[row] == gene2[row]:

            d_score.append(20*risk2[row] + 60)

        elif d_genotype[row] == gene3[row]:

            d_score.append(20*risk3[row] + 80)

    print(AG_num + ' finish calculate "d_score". ' +'\n')



    # D result

    #############
    # (1) 細胞衰老

    # 1.細胞循環

    g1_1 = mean(d_score[0:5])

    # 2.DNA損傷檢查點

    g1_2 = mean(d_score[5:10])

    # 3.端粒酶活性

    g1_3 = mean(d_score[10:12])

    # 4.DNA修復

    g1_4 = mean(d_score[12:20])

    # 5.粒線體活性

    g1_5 = mean(d_score[20:26])

    # 6.對抗自由基

    g1_6 = mean(d_score[26:32])


    #############
    # (2) 毒物代謝

    # 1.肝臟

    g2_1 = mean(d_score[32:37])

    # 2.肺臟

    g2_2 = mean(d_score[37:40])

    # 3.腸道

    g2_3 = mean(d_score[40:43])

    # 4.腎臟

    g2_4 = mean(d_score[43:47])


    #############
    # (3) 營養代謝

    # 1.血管健康

    g3_1 = mean(d_score[47:51])

    # 2.蛋白質代謝

    g3_2 = mean(d_score[51:57])

    # 3.醣類代謝

    g3_3 = mean(d_score[57:65])

    # 4.脂質代謝

    g3_4 = mean(d_score[65:72])


    #############
    # (4) 免疫調節

    # 1.抗原呈現

    g4_1 = mean(d_score[72:77])

    # 2.免疫訓練

    g4_2 = mean(d_score[77:88])

    # 3.發炎因子

    g4_3 = mean(d_score[88:97])



    #####################
    # Score of 4 category

    # (1) 細胞衰老

    g1 = mean(d_score[0:32])

    # (2) 毒物代謝

    g2 = mean(d_score[32:47])

    # (3) 營養代謝

    g3 = mean(d_score[47:72])

    # (4) 免疫調節

    g4 = mean(d_score[72:97])


    #############################################
    # Pick the bad subcategory from each category


    sub = []

    # (1) 細胞衰老

    if min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_1:

        sub.append("細胞循環")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_2:

        sub.append("DNA損傷檢查點")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_3:

        sub.append("端粒酶活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_4:

        sub.append("DNA修復")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_5:

        sub.append("粒線體活性")

    elif min(g1_1, g1_2, g1_3, g1_4, g1_5, g1_6) == g1_6:

        sub.append("對抗自由基")


    # (2) 毒物代謝

    if min(g2_1, g2_2, g2_3, g2_4) == g2_1:

        sub.append("肝臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_2:

        sub.append("肺臟")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_3:

        sub.append("腸道")

    elif min(g2_1, g2_2, g2_3, g2_4) == g2_4:

        sub.append("腎臟")


    # (3) 營養代謝

    if min(g3_1, g3_2, g3_3, g3_4) == g3_1:

        sub.append("血管健康")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_2:

        sub.append("蛋白質代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_3:

        sub.append("醣類代謝")

    elif min(g3_1, g3_2, g3_3, g3_4) == g3_4:

        sub.append("脂質代謝")


    # (4) 免疫調節

    if min(g4_1, g4_2, g4_3) == g4_1:

        sub.append("抗原呈現")

    elif min(g4_1, g4_2, g4_3) == g4_2:

        sub.append("免疫訓練")

    elif min(g4_1, g4_2, g4_3) == g4_3:

        sub.append("發炎因子")


    print(AG_num + ' finish calculate "d_result". ' +'\n')

    sub


    #####################
    # output d_result.csv

    ag_genotype['d_value'] = d_ratio
    ag_genotype['d_genotype'] = d_genotype
    ag_genotype['d_score'] = d_score

    ag_genotype


    outname = AG_num + '_d_result.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    ag_genotype.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')


    #######################
    # output d_category.csv

    category = ["細胞衰老", "毒物代謝", "營養代謝", "免疫調節"]

    d_mean = [format(g1, '.0f'), format(g2, '.0f'), format(g3, '.0f'), format(g4, '.0f')]

    d_category = pd.DataFrame(list(zip(category, d_mean, sub)),
                  columns=['category','mean', 'high risk subcategory'])


    outname = AG_num + '_d_category.csv'

    outdir = '../result/' + AG_sample[2:8] + '/' + AG_sample + '/'
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    fullname = os.path.join(outdir, outname)

    d_category.to_csv(fullname)

    print(AG_num + ' finish output ' + fullname + '\n')


print("---------------------------")
print('細胞衰老: ', format(g1, '.0f'), "分 ", sub[0])
print('毒物代謝: ', format(g2, '.0f'), "分 ", sub[1])
print('營養代謝: ', format(g3, '.0f'), "分 ", sub[2])
print('免疫調節: ', format(g4, '.0f'), "分 ", sub[3])
print("---------------------------")
print("")

print('Congratulation! '+ AG_sample + 'finish all the calculation process from step (1) to (12)!!\n')
