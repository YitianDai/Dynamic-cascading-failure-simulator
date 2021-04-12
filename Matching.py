def MatchingLine_IEEE39(num1, num2):
    import Excel
    value  = '0'    
    data = Excel.read_excel('IEEE39_Parameters.xlsx', 'Line Parameters', 3, 48, 2, 4, 1.0)
    if num1 == '01' and num2 == '01' :
        value = '1'
    if num1 == '02' and num2 == '39' :
        value = '2'
    if num1 == '39' and num2 == '02' :
        value = '2'
    if num1 == '17' and num2 == '17' :
        value = '30'
    if num1 == '18' and num2 == '27' :
        value = '31'
    if num1 == '27' and num2 == '18' :
        value = '31'
    else :
        for i in range(data.shape[1]):
            if num1 == data[1, i] :
                if num2 == data[2, i] : 
                    value = data[0, i]
                    break
            if num1 == data[2, i] :
                if num2 == data[1, i] :
                    value = data[0, i]
                    break
    return value

# # test
# print(MatchingLine_IEEE39('02', '39'))

def MatchingGen_IEEE39(num):
    import Excel
    data = Excel.read_excel('IEEE39_Parameters.xlsx', 'Generator Parameters', 3, 12, 1, 2, 1.0)
    for i in range(data.shape[1]):
        if num == data[0, i] :
            value = data[1, i]
            break
        if num == data[1, i] :
            value = data[0, i]
            break
    return value    


# # test
# print(MatchingGen('02'))

def MatchingLine_ACTIVSg200(num):
    import Excel
    value  = '0'    
    data = Excel.read_excel('ACTIVSg200_Parameters - copy.xlsx','Line Parameters', 3, 247, 1, 16, 1.0)
    for i in range(data.shape[1]):
        if num == data[1, i] :
            value = data[0, i]
            prob = float(data[15, i])
            break
        if num == data[0, i] :
            value = data[1, i]
            prob = float(data[15, i])
            break

    return value, prob 

# # test
# print(MatchingLine_ACTIVSg200('1'))

def MatchingGen_ACTIVSg200(num):
    import Excel
    data = Excel.read_excel('ACTIVSg200_Parameters - copy.xlsx', 'Generator Parameters', 3, 51, 1, 2, 1.0)
    for i in range(data.shape[1]):
        if num == data[0, i] :
            value = data[1, i]
            break
        if num == data[1, i] :
            value = data[0, i]           
            break
    return value   


# # # test
# print(MatchingGen_ACTIVSg200('sym_104_1'))

def MatchingLine_ACTIVSg2000(num):
    import Excel
    value  = '0'    
    data = Excel.read_excel('ACTIVSg2000_Parameters.xlsx','Line Parameters', 3, 3208, 1, 12, 1.0)
    for i in range(data.shape[1]):
        if num == data[1, i] :
            value = data[0, i]
            prob = float(data[11, i])
            break
        if num == data[0, i] :
            value = data[1, i]
            prob = float(data[11, i])
            break

    return value, prob

# # test
# print(MatchingLine_ACTIVSg2000('lne_100_174_1'))

def MatchingGen_ACTIVSg2000(num):
    import Excel
    data = Excel.read_excel('ACTIVSg2000_Parameters.xlsx', 'Generator Parameters', 3, 546, 1, 2, 1.0)
    for i in range(data.shape[1]):
        if num == data[0, i] :
            value = data[1, i]
            break
        if num == data[1, i] :
            value = data[0, i]
            break
    return value    


# # test
# print(MatchingGen_ACTIVSg200('sym_104_1'))

def Matching():
    import Excel
    import csv
    data = Excel.read_excel('Random_N-3_ACTIVSg2000.xlsx', 'Random_N-3_ACTIVSg2000', 1, 1000, 1, 5, 1.0)
    with open("match.csv", 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        for i in range(data.shape[1]):
            num1 = MatchingLine_ACTIVSg2000(data[1, i])
            num2 = MatchingLine_ACTIVSg2000(data[2, i])
            num3 = MatchingLine_ACTIVSg2000(data[3, i])
            writer.writerow([data[0, i], data[1, i], data[2, i], data[3, i], num1, num2, num3, data[4, i]])

# Matching()

def RC_Matching():
    import Excel
    import csv
    data = Excel.read_excel('RC_ACTIVSg200.xlsx', 'N-2', 1, 784, 6, 7, 1.0)
    

    with open("match.csv", 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        for i in range(data.shape[1]):
            [num1, prob1] = MatchingLine_ACTIVSg200(data[0, i])
            [num2, prob2] = MatchingLine_ACTIVSg200(data[1, i])
            # num3 = MatchingLine_ACTIVSg2000(data[3, i])
            print(i, num1, num2, prob1, prob2)
            writer.writerow([i, num1, num2, data[0, i], data[1, i], float(prob1)*float(prob2)])
    
# RC_Matching()

def delete_none(input):
    new2 = []
    for item in input:
        if item:
            new1 = []
            for x in item:
                if x:
                    new1.append(x)
            new2.append(new1)
    return new2


# print(delete_none(delete_none([[''], [''], ['', '3193'], ['', '3166', '3056', '3089'], ['', '3014', '3122'], [''], [''], [''], [''], ['']])))