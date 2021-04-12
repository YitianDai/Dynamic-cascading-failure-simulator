# with generator tripping
def events_recorder(input, name, ShowFullMessages):
    import csv
    import re
    import Excel
    import Matching
    import xlrd
    import numpy as np
    import Matlab

    p1 = re.compile(r'[=](.*?)[)]', re.S)
    p2 = re.compile(r"Grid split into (.*)", re.S)
    p3 = re.compile(r"Grid\\(.*).ElmLne", re.S)
    p4 = re.compile(r"Grid\\Line(.*)[-](.*).ElmLne", re.S)
    p5 = re.compile(r"Grid\\Bus (.*)\\Cub", re.S)
    p6 = re.compile(r"Element (.*) is local reference", re.S)
    p7 = re.compile(r"  1 (.*)", re.S)
    p8 = re.compile(r"Grid\\(.*).ElmSym", re.S)
    p9 = re.compile(r"Step: (.*)[)]", re.S)
    p10 = re.compile(r"\\Cub_(.*)\\Switch", re.S)
    p11 = re.compile(r"G (.*)", re.S)


    # Global variables
    count = 0
    flag = 0
    flag2 = 3
    Cub1 = ""
    Comp1 = ""
    Name_of_Local_Reference2 = ""
    Switch_Event1 = ""
    Unsupplied_Areas1 = ""
    State_of_Logic1 = ""
    Targeted_Load2 = ""
    Percent_of_Shedding1 = ""
    Amount_of_Load_Shedding = 0.0
    Amount_of_Load_Shedding1 = 0.0
    Amount_of_Load_Shedding2 = 0.0
    faulty_line = []
    faulty_node_overfreq = []
    faulty_node_underfreq = []
    tripped_line = []
    tripped_node_overfreq = []
    tripped_node_underfreq = []
    tripped_load = []
    tripped_LS_balance = []
    tripped_LS_tripped_gen = []
    tripped_LS_tripped_load = []
    tripped_LS_unsupplied = []
    failed_matrix = []
    faulty_matrix = []
    OutOfStep = []
    State_of_generator = ""
    amount_of_shedding= [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    j = 0
    i = 0

    # Active Power of Loads
    Loads = Excel.read_excel('IEEE39_Parameters.xlsx', 'Load Parameters', 8, 26, 1, 3, 1.0)
    target_load = Loads[0,:]

    # create empty matrix of 512*1
    data = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], ['']]

    with open(name, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        f = open(input, 'r')
        for line in f.readlines():
            temp = line.strip()
            if temp.find('.ElmSym') >= 0:
                Comp = re.findall(p8, temp)
                Comp1 = "".join(Comp)
            elif temp.find('Generator out of step (pole slip)') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                data[i].extend([t1, Comp1, 'Out of step', '', ''])
                i = i + 1
                Comp1 = ""
            elif temp.find('Outage Event: Element set to out of service') >= 0:
                Name_of_Local_Reference2 = ""
                Unsupplied_Areas1 = ""
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Trip_of_generator = ""
                if Comp1 != "" :
                    if Comp1.find("(") == -1 :
                        gen = re.findall(p11, Comp1)
                        gen1 = "".join(gen)
                        name = Comp1 
                        State_of_generator = '>360 tripped'
                        Trip_of_generator = Comp1 + ' All tripped'
                        faulty_node_overfreq.append(Matching.MatchingGen_IEEE39(gen1))
                    else :
                        name = Comp1 
                        State_of_generator = '>360 tripped'
                        Trip_of_generator = ""
                    data[i].extend([t1, name, State_of_generator, Trip_of_generator, ''])
                    i = i + 1
                    Comp1 = ""
            elif temp.find('.ElmLne') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Comp1 = ""
                if temp.find('evt  -') >= 0:
                    Line_Outage = re.findall(p3, temp)
                    Line_Outage1 = "".join(Line_Outage)
                    Switch_Event = re.findall(p4, temp)
                    Switch_Event1 = " ".join(Switch_Event[0])
                    Switch_Event2 = Switch_Event1.split(' ')
                    row = Matching.MatchingLine_IEEE39(Switch_Event2[0], Switch_Event2[1])
                    faulty_line.append(row)
                    data[i].extend([t1, Line_Outage1, row, '', ''])
                    i = i + 1
            
            elif temp.find('.StaSwitch') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                if flag == 0:
                    Switch_Event = re.findall(p5, temp)
                    Switch_Event1 = "".join(Switch_Event)
                    Cub = re.findall(p10, temp)
                    Cub1 = "".join(Cub)
                    if (Switch_Event1 == '30') or (Switch_Event1 == '32') or (Switch_Event1 == '33') or (Switch_Event1 == '34') or (Switch_Event1 == '35') or (Switch_Event1 == '36') or (Switch_Event1 == '37') or (Switch_Event1 == '38') :
                        if (Cub1 == '1') or (Cub1 == '1(1)') or (Cub1 == '2(1)') or (Cub1 == '3'):
                            if Cub1 == '1' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 4'
                                    faulty_node_overfreq.append(Switch_Event1)
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 4'
                                    faulty_node_underfreq.append(Switch_Event1)
                                Trip_of_generator = 'Gen' + Switch_Event1 + ' tripped'
                                
                            elif Cub1 == '1(1)' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 3'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 3'
                                Trip_of_generator = ""
                            elif Cub1 == '2(1)' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 2'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 2'
                                Trip_of_generator = ""
                            elif Cub1 == '3':
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 1'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 1'
                                Trip_of_generator = ""
                            if ShowFullMessages:
                                data[i].extend([t1, 'Gen' + Switch_Event1, State_of_generator, Trip_of_generator, ''])
                                i = i + 1
                        else:
                            flag = 1
                    elif (Switch_Event1 == '31') or (Switch_Event1 == '39'):
                        if (Cub1 == '1') or (Cub1 == '1(1)') or (Cub1 == '2(1)') or (Cub1 == '3(1)'):
                            if Cub1 == '1' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 4'
                                    faulty_node_overfreq.append(Switch_Event1)
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 4'
                                    faulty_node_underfreq.append(Switch_Event1)
                                Trip_of_generator = 'Gen' + Switch_Event1 + ' tripped'
                                
                            elif Cub1 == '1(1)' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 3'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 3'
                                Trip_of_generator = ""
                            elif Cub1 == '2(1)' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 2'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 2'
                                Trip_of_generator = ""
                            elif Cub1 == '3(1)' :
                                if flag2 == 0:
                                    State_of_generator = '>F Trip 1'
                                elif flag2 == 1:
                                    State_of_generator = '<F Trip 1'
                                Trip_of_generator = ""
                            if ShowFullMessages:
                                data[i].extend([t1, 'Gen' + Switch_Event1, State_of_generator, Trip_of_generator, ''])
                                i = i + 1
                        else:
                            flag = 1
                    else:
                        flag = 1
                elif flag == 1:
                    Switch_Event = re.findall(p5, temp)
                    Switch_Event2 = "".join(Switch_Event)
                    if flag2 == 2 : 
                        Line_Outage = "Trf" + Switch_Event1 + "-" + Switch_Event2
                    else :
                        Line_Outage = "Line" + Switch_Event1 + "-" + Switch_Event2
                    row = Matching.MatchingLine_IEEE39(Switch_Event1, Switch_Event2)
                    faulty_line.append(row)
                    data[i].extend([t1, Line_Outage, row, '', ''])
                    i = i + 1
                    flag = 0
            elif temp.find('local reference') >= 0:
                Name_of_Local_Reference = re.findall(p6, temp)
                Name_of_Local_Reference1 = "".join(Name_of_Local_Reference)
                Name_of_Local_Reference3 = Name_of_Local_Reference1 + Name_of_Local_Reference2
                Name_of_Local_Reference2 = Name_of_Local_Reference3
            elif temp.find('area(s) are unsupplied') >= 0:
                Unsupplied_Areas = re.findall(p7, temp)
                Unsupplied_Areas1 = "".join(Unsupplied_Areas)
            elif temp.find('Grid split') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                No_of_Islands = re.findall(p2, temp)
                No_of_Islands1 = "".join(No_of_Islands)
                if ShowFullMessages:
                    data[i].extend([t1, No_of_Islands1, Name_of_Local_Reference2, Unsupplied_Areas1, ''])
                    i = i + 1
            elif temp.find('Circuit-Breaker Action') >= 0 :
                Name_of_Local_Reference2 = ""
                Unsupplied_Areas1 = ""
            elif temp.find('.ElmRelay') >= 0:
                if temp.find('Model.ElmRelay') >= 0:
                    flag2 = 0
                elif temp.find('Model(1).ElmRelay') >= 0:
                    flag2 = 1
                elif temp.find('OCTF.ElmRelay') >= 0:
                    flag2 = 2
                Targeted_Load = re.findall(p5, temp)
                Targeted_Load1 = "".join(Targeted_Load)
                Targeted_Load2 = 'Load ' + Targeted_Load1
                for j in range(Loads.shape[1]):
                    if Targeted_Load2 == Loads[1, j]:
                        Amount_of_Load_Shedding = float(Loads[2, j])
                        break
            elif temp.find('Relay is tripping') >= 0:
                Targeted_Load2 = ""
                Amount_of_Load_Shedding = 0.0
                j = 100
            elif temp.find('Frequency Relay') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                State_of_Logic = re.findall(p9, temp)
                State_of_Logic1 = "".join(State_of_Logic)
                if temp.find('1 Logic') >= 0 or temp.find('2 Logic') >= 0 or temp.find('8 Logic') >= 0 or temp.find('9 Logic') >= 0 :
                    Percent_of_Shedding = '-5%'
                    Amount_of_Load_Shedding1 = Amount_of_Load_Shedding * 0.05
                elif temp.find('3 Logic') >= 0:
                    Percent_of_Shedding = '-10%'
                    Amount_of_Load_Shedding1 = Amount_of_Load_Shedding * 0.1
                elif temp.find('4 Logic') >= 0 or temp.find('5 Logic') >= 0 or temp.find('6 Logic') >= 0 or temp.find('7 Logic') >= 0:
                    Percent_of_Shedding = '-7.5%'
                    Amount_of_Load_Shedding1 = Amount_of_Load_Shedding * 0.075
                Percent_of_Shedding1 = "".join(Percent_of_Shedding)
                if temp.find('1 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.05
                elif temp.find('2 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.1
                elif temp.find('3 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.2
                elif temp.find('4 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.275
                elif temp.find('5 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.35
                elif temp.find('6 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.425
                elif temp.find('7 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.5
                elif temp.find('8 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.55
                elif temp.find('9 Logic') >= 0 :
                    amount_of_shedding[j] =  Amount_of_Load_Shedding * 0.6
                if ShowFullMessages:
                    data[i].extend([t1, Targeted_Load2, State_of_Logic1 + Percent_of_Shedding1, Amount_of_Load_Shedding1, ''])
                    i = i + 1
                Amount_of_Load_Shedding3 = Amount_of_Load_Shedding2 + Amount_of_Load_Shedding1
                Amount_of_Load_Shedding2 = Amount_of_Load_Shedding3
            elif temp.find('Simulation successfully executed.') >= 0 or temp.find('System-Matrix Inversion failed') >= 0 :
                print(faulty_line, faulty_node_overfreq, faulty_node_underfreq, amount_of_shedding)
                
                [a, b, c, LS, LS_balance, LS_tripped_gen, LS_tripped_load, LS_unsupplied] = Matlab.Dynamic_Load_Shedding_Calculator_IEEE39(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding)
                data[i].extend([a, b, c, Amount_of_Load_Shedding2, LS])
                i = i + 1
                print("LS = %.10f" % LS)
                print("LS_balance = %.10f" % LS_balance)
                print("LS_tripped_gen = %.10f" % LS_tripped_gen)
                print("LS_tripped_load = %.10f" % LS_tripped_load)
                print("LS_unsupplied = %.10f" % LS_unsupplied)
                tripped_line.append(a)
                tripped_node_overfreq.append(b)
                tripped_node_underfreq.append(c)
                tripped_load.append(LS)
                tripped_LS_balance.append(LS_balance)
                tripped_LS_tripped_gen.append(LS_tripped_gen)
                tripped_LS_tripped_load.append(LS_tripped_load)
                tripped_LS_unsupplied.append(LS_unsupplied)
                if temp.find('System-Matrix Inversion failed') >= 0 :
                    data[i].extend([t1, 'System-Matrix Inversion failed', '', '', ''])
                    i = i + 1
                    faulty_matrix = 'System-Matrix Inversion failed'
                    failed_matrix.append(faulty_matrix)
                else:
                    faulty_matrix = ''
                    failed_matrix.append(faulty_matrix)
                faulty_line = []
                faulty_node_overfreq = []
                faulty_node_underfreq = []
                faulty_matrix = []
                LS = 0
                amount_of_shedding= [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                Amount_of_Load_Shedding2 = 0
                if i <= 512:
                    for i in range(i, 512):
                        data[i].extend(['', '', '', '', ''])
                i = 0       
        linematrix = np.array(tripped_line)
        nodeoverfreqmatrix = np.array(tripped_node_overfreq)
        nodeunderfreqmatrix = np.array(tripped_node_underfreq)
        loadmatrix = np.array(tripped_load)
        loadbalancematrix = np.array(tripped_LS_balance)
        loadtrippedgenmatrix = np.array(tripped_LS_tripped_gen)
        loadtrippedloadmatrix = np.array(tripped_LS_tripped_load)
        loadunsuppliedmatrix = np.array(tripped_LS_unsupplied)
        Matrixfailed = np.array(failed_matrix)
        #print(linematrix, nodeoverfreqmatrix, nodeunderfreqmatrix, loadmatrix)
        for i in range(512):
            d = data[i]
            writer.writerow(d)
    with open('244-560_summary.csv', 'w', newline='') as csvfile:  
        writer = csv.writer(csvfile)
        for m in range(loadmatrix.shape[0]):
            writer.writerow([linematrix[m], nodeoverfreqmatrix[m], nodeunderfreqmatrix[m], loadmatrix[m], loadbalancematrix[m], loadtrippedgenmatrix[m], loadtrippedloadmatrix[m], loadunsuppliedmatrix[m], Matrixfailed[m]])


#events_recorder('N-2_100loading_N-1_Unsecure.txt', 'N-2_100loading_N-1_Unsecure_Shortcut.csv', 0)
# events_recorder('1.txt', "1.csv", 1)
events_recorder('244-560.txt', "244-560.csv", 1)
# events_recorder('N-2_General.txt', "N-2_General.csv", 1)
