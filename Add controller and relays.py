
import powerfactory as pf
app = pf.GetApplication()   

# #get all objects/elements within the network
# Lines = app.GetCalcRelevantObjects("*.ElmLne")
# Buses = app.GetCalcRelevantObjects("*.ElmTerm")
# Loads = app.GetCalcRelevantObjects("*.ElmLod")
# Cubs = app.GetCalcRelevantObjects("*.StaCubic")
# Switches = app.GetCalcRelevantObjects("*.StaSwitch")
# Relays = app.GetCalcRelevantObjects("*.ElmRelay")
# Transformers = app.GetCalcRelevantObjects("*.ElmTr2")


# change the model of the synchronous generator
def ChangeSymType(type) :
    sMachineTypes = app.GetCalcRelevantObjects("*.TypSym")
    for item in sMachineTypes:
        if type == 'standard' : 
            item.model_inp='det'
        if type == 'classical' : 
            item.model_inp='cls'


def ChangeLoadType() :
    LoadTypes = app.GetCalcRelevantObjects("*.TypLod")
    
    for item in LoadTypes:
        # app.PrintPlain(item)
        item.aP = 0.391
        item.bP = 0.42
        item.bQ = -1


def DisableAVR() :
    AVR = app.GetCalcRelevantObjects("*.ElmDsl")

    for item in AVR:
        if item.loc_name.find('AVR') >= 0 or item.loc_name.find('SEXS') >= 0 or item.loc_name.find('IEEEVC') >= 0:
            # app.PrintPlain(item)
            item.outserv = 1


def AddOvercurrentRelay(items = None):
    SelectedLines = []
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    Transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    if items == None:
        SelectedLines = Lines
    else:
        for item in items:
            for line in Lines:
                if line.loc_name == item:
                    SelectedLines.append(line)
                    break
    # app.PrintPlain(SelectedLines)
    # app.PrintPlain(len(SelectedLines))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    OvercurrentRelay = RelayFolder.GetContents('Overcurrent Relay.TypRelay')[0]
    CT = OvercurrentRelay.GetContents('*.TypCt')
    # app.PrintPlain(OvercurrentRelay)

    import Excel
    data = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Line Parameters', 3, 247, 2, 12, 1.0)

    for line in SelectedLines:
        cub1 = line.bus1
        cub2 = line.bus2
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        # app.PrintPlain(ExistRelay)
        for relay in ExistRelay:
            if relay.loc_name == 'Overcurrent Relay':
                app.PrintPlain([line, r'Overcurrent Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', 'Overcurrent Relay')
        relay.typ_id = OvercurrentRelay
        ct = relay.CreateObject('StaCt', 'CT')
        corect = relay.CreateObject('StaCt', 'CoreCT')
        ct.typ_id = CT[0]
        ct.ptapset = 1000
        corect.typ_id = CT[0]
        corect.ptapset = 1000
        relay.slotupd()


        P11 = relay.pdiselm[4]
        P01 = relay.pdiselm[5]
        P02 = relay.pdiselm[6]
        for i in range(data.shape[1]):
            if line.loc_name == data[0, i]:
                P11.Ipsetr = float(data[9, i]) * 1.35
                P11.Tpset = float(data[10, i])
                P01.Ipsetr = float(data[5, i]) * 1.35
                P01.Tset = float(data[6, i])
                P02.Ipsetr = float(data[7, i])  * 1.35
                P02.Tset = float(data[8, i]) * 1.0
                break
        
        
        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        for item in Logics:
            if item.fold_id.loc_name == 'Overcurrent Relay':
                if item.pSwitch == [None]:
                    item.pSwitch=[cub1, cub2]
        app.PrintPlain([line, r'Overcurrent Relay Installed'])

    for Trf in Transformers:
        cub1 = Trf.bushv
        cub2 = Trf.buslv
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name == 'Overcurrent TF Relay':
                app.PrintPlain([Trf, r'Overcurrent TF Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', 'Overcurrent TF Relay')
        relay.typ_id = OvercurrentRelay
        ct = relay.CreateObject('StaCt', 'CT')
        corect = relay.CreateObject('StaCt', 'CoreCT')
        ct.typ_id = CT[0]
        ct.ptapset = 1000
        corect.typ_id = CT[0]
        corect.ptapset = 1000
        relay.slotupd()


        P11 = relay.pdiselm[4]
        P01 = relay.pdiselm[5]
        P02 = relay.pdiselm[6]
        for i in range(data.shape[1]):
            if Trf.loc_name == data[0, i]:
                P11.Ipsetr = float(data[9, i]) * 1.0
                P11.Tpset = float(data[10, i])
                P01.Ipsetr = float(data[5, i])
                P01.Tset = float(data[6, i])
                P02.Ipsetr = float(data[7, i])  * 1.35
                P02.Tset = float(data[8, i]) * 1.0
                break
        
        
        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        for item in Logics:
            if item.fold_id.loc_name == 'Overcurrent TF Relay':
                if item.pSwitch == [None]:
                    item.pSwitch=[cub1, cub2]
        app.PrintPlain([Trf, r'Overcurrent TF Relay Installed'])



def RemoveOvercurrentRelay(items = None):
    SelectedLines = []
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    Transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    if items == None:
        SelectedLines = Lines
    else:
        for item in items:
            for line in Lines:
                if line.loc_name == item:
                    SelectedLines.append(line)
                    break
    # app.PrintPlain(SelectedLines)
    # app.PrintPlain(len(SelectedLines))
    for line in SelectedLines:
        cub1 = line.bus1
        cub2 = line.bus2
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name == 'Overcurrent Relay':
                app.PrintPlain([line, r'Overcurrent Relay Removed'])
                item.Delete()


    for Trf in Transformers:
        cub1 = Trf.bushv
        cub2 = Trf.buslv
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name == 'Overcurrent TF Relay':
                app.PrintPlain([Trf, r'Overcurrent TF Relay Removed'])
                item.Delete()


def AddUnderFrequencyLoadShedding(items = None):
    SelectedLoads = []
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    if items == None:
        SelectedLoads = Loads
    else:
        for item in items:
            for load in Loads:
                if load.loc_name == item:
                    SelectedLoads.append(load)
                    break
    # app.PrintPlain(SelectedLoads)
    # app.PrintPlain(len(SelectedLoads))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    UnderFrequencyLoadShedding = RelayFolder.GetContents('LFDD.TypRelay')[0]


    for load in SelectedLoads:
        cub1 = load.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name == 'UFLS Relay':
                app.PrintPlain([load, r'UFLS Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', 'UFLS Relay')
        relay.typ_id = UnderFrequencyLoadShedding
        relay.slotupd()

        Logics = app.GetCalcRelevantObjects("*.RelLslogic")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name == 'UFLS Relay' and item.pLoad == [None]:
                    # app.PrintPlain('asd')
                    item.pLoad = [load]

        PLLs = app.GetCalcRelevantObjects("*.ElmPhi__pll")
        # app.PrintPlain(PLLs)
        for item in PLLs:
            if item.fold_id.loc_name == 'UFLS Relay' and item.pbusbar == None:
                item.pbusbar = load.bus1
                item.mversion = 2


def RemoveUnderFrequencyLoadShedding(items = None):
    SelectedLoads = []
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    if items == None:
        SelectedLoads = Loads
    else:
        for item in items:
            for load in Loads:
                if load.loc_name == item:
                    SelectedLoads.append(load)
                    break
    # app.PrintPlain(SelectedLoads)
    # app.PrintPlain(len(SelectedLoads))
    for load in SelectedLoads:
        cub1 = load.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name == 'UFLS Relay':
                app.PrintPlain([load, r'UFLS Relay Removed'])
                item.Delete()


def AddOverFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    # app.PrintPlain(SelectedGens)
    # app.PrintPlain(len(SelectedGens))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    OverFrequencyGeneratorTripping = RelayFolder.GetContents('OFGT.TypRelay')[0]
    # app.PrintPlain(TypRelays)

    for Gen in SelectedGens:
        cub1 = Gen.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name == 'OFGT Relay':
                app.PrintPlain([Gen, r'OFGT Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', 'OFGT Relay')
        relay.typ_id = OverFrequencyGeneratorTripping
        relay.slotupd()

        Logics = app.GetCalcRelevantObjects("*.RelLogic")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name == 'OFGT Relay':                
                if item.pSwitch == [None]:
                    item.pSwitch = [cub1]


def RemoveOverFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    for Gen in SelectedGens:
        cub1 = Gen.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name == 'OFGT Relay':
                app.PrintPlain([Gen, r'OFGT Relay Removed'])
                item.Delete()


def AddUnderFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    # app.PrintPlain(SelectedGens)
    # app.PrintPlain(len(SelectedGens))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    UnderFrequencyGeneratorTripping = RelayFolder.GetContents('UFGT.TypRelay')[0]
    # app.PrintPlain(TypRelays)

    for Gen in SelectedGens:
        cub1 = Gen.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name == 'UFGT Relay':
                app.PrintPlain([Gen, r'UFGT Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', 'UFGT Relay')
        relay.typ_id = UnderFrequencyGeneratorTripping
        relay.slotupd()

        Logics = app.GetCalcRelevantObjects("*.RelLogic")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name == 'UFGT Relay':                
                if item.pSwitch == [None]:
                    item.pSwitch = [cub1]
        


def RemoveUnderFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    for Gen in SelectedGens:
        cub1 = Gen.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name == 'UFGT Relay':
                app.PrintPlain([Gen, r'UFGT Relay Removed'])
                item.Delete()




def SetUpGeneratorControl():
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    UserDefinedModels = app.GetProjectFolder('blk')
    GenControl = UserDefinedModels.GetContents('SYM Frame.BlkDef')[0]
    AGC = UserDefinedModels.GetContents('AGC.BlkDef')[0]
    OutOfStep = UserDefinedModels.GetContents('OutOfStep.BlkDef')[0]
    # app.PrintPlain(UserDefinedModels)
    # app.PrintPlain(GenControl)

    for SymMachine in SymMachines:
        CompModel = SymMachine.c_pmod
        CompModel.typ_id = GenControl
        ExistDsl = CompModel.GetContents('*.ElmDsl')
        for item in ExistDsl:
            if item.loc_name.find('AGC') >= 0 :
                app.PrintPlain([item, r'AGC Exists'])
                item.Delete()
            if item.loc_name.find('OutOfStep') >= 0 :
                app.PrintPlain([item, r'OutOfStep Exists'])
                item.Delete()
        
        ExistVmea = CompModel.GetContents('*.StaVmea')
        for item in ExistVmea:
            if item.loc_name.find('Vmea') >= 0 :
                app.PrintPlain([item, r'Vmea Exists'])
                item.Delete()
            
        AGCDsl = CompModel.CreateObject('ElmDsl', 'AGC')
        AGCDsl.typ_id = AGC
        AGCDsl.params = [0.002, 0.1]
        OutOfStepDsl = CompModel.CreateObject('ElmDsl', 'OutOfStep')
        OutOfStepDsl.typ_id = OutOfStep
        CompModel.SlotUpdate()
        Vmea1 = CompModel.CreateObject('StaVmea', 'Vmea1')
        Vmea1.pbusbar = SymMachine.bus1
        Vmea2 = CompModel.CreateObject('StaVmea', 'Vmea2')
        Vmea2.pbusbar = SymMachine.bus1

        CompModel.SlotUpdate()

        pelm = CompModel.pelm
        pelm[0] = SymMachine
        pelm[11] = SymMachine

        CompModel.SetAttribute('pelm', pelm)
        app.PrintPlain(CompModel.pelm)


def ChangeLoadingLevel(loading = 1.0):
    import Excel
    
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    Ldf = app.GetFromStudyCase('ComLdf')
    data = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Load Parameters', 8, 167, 2, 4, loading)
    app.PrintPlain(loading)
    for Load in Loads:
        for i in range(data.shape[1]):
            if Load.loc_name == data[0, i]:
                Load.plini = float(data[1, i])
                Load.qlini = float(data[2, i])
                break
    # Ldf.Execute()

def ChangeInitialGenerationLevel(loading = 1.0):
    import Excel
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    data = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Pmin=0 reserve=60', 4, 52, 3, 9, 1.0)
    Ldf = app.GetFromStudyCase('ComLdf')
    
    for item in SymMachines:
        for i in range(data.shape[1]):
            if item.loc_name == data[0, i]:
                # app.PrintPlain(item.loc_name)
                if float(data[3, i]) == 1.0 :
                    item.outserv = 0
                    # app.PrintPlain(item.loc_name)
                    
                else:
                    item.outserv = 1
                item.pgini = float(data[4, i])
                item.qgini = float(data[5, i])
                # app.PrintPlain(float(data[6, i]))
                item.usetp = float(data[6, i])
                item.Pmin_uc = float(data[2, i]) * loading
                # item.pgini = float(data[2, i]) * 1.0
                # item.qgini = float(data[3, i]) * 1.0
                # Load.qlini = 0.0
                break
    Ldf.Execute()

def ChangeLineRating():
    
    import Excel
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    name = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Line Parameters', 3, 181, 2, 2, 1.0)
    rating = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Line Parameters', 3, 181, 9, 9, 1.0)

    for line in Lines:
        for i in range(name.shape[1]):
            if line.loc_name == name[0, i]:
                line.typ_id.sline = float(rating[0, i]) * 1.35
                break


def GenerateRandomCases(num):
    import itertools
    import xlwt
    import csv
    import random
    import Matching
    import numpy as np
    import Excel

    cases = []
    comb = []
    combs = []
    prob1 = 0
    prob2 = 0
    prob = []
    data = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Line Parameters', 3, 181, 2, 15, 1.0)

    for lines in itertools.combinations(Lines, 2) :
        comb.append(lines[0])
        comb.append(lines[1])
        combs.append(comb)
        for i in range(data.shape[1]):
            if lines[0].loc_name == data[0, i] :
                prob1 = float(data[13, i])
            if lines[1].loc_name == data[0, i] :
                prob2 = float(data[13, i])
        if prob1*prob2 == 0:
            app.PrintPlain('Probability calculation ERROR')
        
        prob.append(prob1*prob2)
        # app.PrintPlain(prob1*prob2)
        comb= []
        prob1 = 0
        prob2 = 0

    # randomly choose
    # cases = random.sample(combs, num)

    # weighted choose 
    app.PrintPlain(prob)
    cases = [weighting_choose(combs, prob) for _ in range(num)]
    app.PrintPlain(cases)

    with open("Random Cases.csv", 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        for i in range(len(cases)):
            for item in Lines:
                if item == cases[i][0]:
                        a = item.loc_name
                if item == cases[i][1]:
                        b = item.loc_name
            writer.writerow([i, a, b, prob[i]])


    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'Random Cases',cell_overwrite_ok=True)

    for i in range(len(cases)):
        sheet1.write(i, 0, i)
        sheet1.write(i, 1, str(cases[i][0].loc_name))
        sheet1.write(i, 2, str(cases[i][1].loc_name))
        sheet1.write(i, 3, str(prob[i]))
    
    f.save('Random Cases.xlsx')


def RunN2Contingencies():
    
    import time
    import os
    import csv
    import pandas as pd

    time_start = time.time()

    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    Ldf = app.GetFromStudyCase('ComLdf')   # Get commands of calculating load flow
    Init = app.GetFromStudyCase('ComInc')  # Get commands of calculating initial conditions
    Sim = app.GetFromStudyCase('ComSim')   # Get commands of running simulations
    ElmRes = app.GetFromStudyCase('Results.ElmRes')  # Create class of result variables named "Results"
    ComRes = app.GetFromStudyCase('ComRes')  # Get commands of export results
    Events_folder = app.GetFromStudyCase('IntEvt')  # Get events folder

    start = 0

    end = 1000
    
    EventSet = Events_folder.GetContents()
    Outage1 = EventSet[0]
    Outage2 = EventSet[1]
    Outage1.outserv = 0
    Outage2.outserv = 0 # switch on the outage event


    data = pd.read_csv("Random Cases.csv", header=None)
    list=data.values.tolist()
    # app.PrintPlain(list[0])
    # app.PrintPlain(list[0][1])
    # app.PrintPlain(len(list))


    for i in range(start, end) :

        app.PrintPlain(i)
        app.PrintPlain(list[i][1:3])


        # Init.Execute()
        #lines1 = np.array(lines)

        for item in Lines:
            if item.loc_name == list[i][1]:
                Outage1.p_target = item # event target is the each line
                Outage1.time = 1.0  # starts at t= 1s
                Outage1.i_what = 0  # take the element out of service
            if item.loc_name == list[i][2]:
                Outage2.p_target = item
                Outage2.time = 1.0  # starts at t= 1s
                Outage2.i_what = 0  # take the element out of service

        Sim.tstop = 300 # simulation time = 300s

        Init.Execute()
        Sim.Execute()
        
        Ldf.Execute()

        time_end = time.time()
        # with open("readme2.csv", 'a', newline='') as csvfile:
        #     writer = csv.writer(csvfile)
        #     writer.writerow([i, list[i][1:3], time_end - time_start])
        Window = app.GetOutputWindow()

        Path= r'C:\Users\DayDay\OneDrive - University of Manchester\Yitian\Matpower\0.65\%i-%i.txt'%(start, i)
        # Path= r'C:\Users\44756\OneDrive - UOM\Yitian\Matpower\%i-%i.txt'%(start, i)
        

        Window.Save(Path)
        #app.ClearOutputWindow()
        
        if (i - start) > 0 : 
            Path_old= r'C:\Users\DayDay\OneDrive - University of Manchester\Yitian\Matpower\0.65\%i-%i.txt'%(start, i - 1)
            # Path_old= r'C:\Users\44756\OneDrive - UOM\Yitian\Matpower\%i-%i.txt'%(start, i - 1)
            os.remove(Path_old)

def matching():
    import Excel
    import csv
    import math 
    import xlwt

    # ################### output specific parameters ####################################
    # with open("match.csv", 'a', newline='') as csvfile:
    #     writer = csv.writer(csvfile)

    #     for item in Transformers:
    #         writer.writerow([item.typ_id, item.typ_id.utrn_h])


    staticdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Pmin=0 reserve=60 (2)', 3, 203, 6, 28, 1.0)
    dynamicdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Pmin=0 reserve=60 (2)', 3, 52, 2, 2, 1.0)

    with open("match.csv", 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        for i in range(dynamicdata.shape[1]):
            for j in range(staticdata.shape[1]):

            
                # print(staticdata[0, j])
                # print(dynamicdata[0, i])
                if staticdata[0, j] == dynamicdata[0, i]:
                    # print(staticdata.shape[0])
                    # writer.writerow([dynamicdata[0, i], dynamicdata[15, i], staticdata[0, j]])
                    writer.writerow(staticdata[0: staticdata.shape[0], j])
                    # writer.writerow([staticdata[0, j], staticdata[1, j], staticdata[2, j], math.sqrt(3)*float(dynamicdata[8, i])*float(dynamicdata[13, i])])
                    break

                # if staticdata[0, j] != dynamicdata[0, i] and i == dynamicdata.shape[1] -1:
                #     writer.writerow([staticdata[0, j], staticdata[1, j], staticdata[2, j], staticdata[3, j], 'no'])


    # ########### get the correlation between digsilent and matpower ##########################################
    # bus1 = []
    # bus2 = []
    # bus3 = []

    # staticdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'OPF PQV Pmin=0', 4, 52, 5, 45, 1.0)
    # dynamicdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'OPF PQV Pmin=0', 4, 52, 1, 1, 1.0)
    
    
    # for i in range(dynamicdata.shape[1]):
    # ###### for line ###########
    #     # item = re.findall(re.compile(r"lne_(.*)_(.*)_(.*)", re.S), dynamicdata[0, i])
    # ###### for transformer ###########
    #     item = re.findall(re.compile(r"trf_(.*)_(.*)_(.*)", re.S), dynamicdata[0, i])
    #     item1 = " ".join(item[0])
    #     item2 = item1.split(' ')
    #     bus1.append(item2[0])
    #     bus2.append(item2[1])
    #     bus3.append(item2[2])
    
    
    # with open("match1.csv", 'a', newline='') as csvfile:
    #     writer = csv.writer(csvfile)
    #     for i in range(dynamicdata.shape[1]):
    #         for j in range(staticdata.shape[1]):
    #             # print(staticdata[1, j])
    #             # print(bus1[i])
    #             if staticdata[1, j] == bus1[i] and staticdata[2, j] == bus2[i]:
    #                 # writer.writerow([dynamicdata[0, i], float(staticdata[0, j]) + float(bus3[i]) - 1, staticdata[1, j], staticdata[2, j], staticdata[3, j], math.sqrt(3)*float(dynamicdata[1, i])*float(dynamicdata[9, i])])
    #                 writer.writerow([dynamicdata[0, i], float(staticdata[0, j]) + float(bus3[i]) - 1, staticdata[1, j], staticdata[2, j]])
    #                 break
    #             if staticdata[2, j] == bus1[i] and staticdata[1, j] == bus2[i]:
    #                 # writer.writerow([dynamicdata[0, i], float(staticdata[0, j]) + float(bus3[i]) - 1, staticdata[1, j], staticdata[2, j], staticdata[3, j], math.sqrt(3)*float(dynamicdata[1, i])*float(dynamicdata[9, i])])
    #                 writer.writerow([dynamicdata[0, i], float(staticdata[0, j]) + float(bus3[i]) - 1, staticdata[1, j], staticdata[2, j]])
    #                 break

    #     ########### get the correlation between matpower and digsilent (generator) ######################################
    # staticdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Generator Parameters', 3, 51, 13, 14, 1.0)
    # dynamicdata = Excel.read_excel('ACTIVSg200_Parameters.xlsx', 'Generator Parameters', 3, 546, 2, 2, 1.0)


    # bus1 = []
    # bus2 = []

    # for i in range(dynamicdata.shape[1]):
    #     item = re.findall(re.compile(r"sym_(.*)_(.*)", re.S), dynamicdata[0, i])
    #     item1 = " ".join(item[0])
    #     item2 = item1.split(' ')
    #     bus1.append(item2[0])
    #     bus2.append(item2[1])


    # with open("match2.csv", 'a', newline='') as csvfile:
    #     writer = csv.writer(csvfile)
    #     for i in range(dynamicdata.shape[1]):
    #         for j in range(staticdata.shape[1]):
    #             # print(staticdata[1, j])
    #             # print(bus1[i])
    #             if staticdata[1, j] == bus1[i]:
    #                 writer.writerow([dynamicdata[0, i], float(staticdata[0, j]) + float(bus2[i]) - 1, staticdata[1, j]])
    #                 print('yes')
    #                 break
                    
                    

def __in_which_part(n, w):
    for i, v in enumerate(w):
        if n < v:
            print("n = ", n) 
            return i
    return len(w) - 1
 
def weighting_choose(data, weightings):
    import random
    s = sum(weightings)
    w = [float(x)/s for x in weightings]
    t = 0
    for i, v in enumerate(w):
        t += v
        w[i] = t
    # app.PrintPlain(w)
    c = __in_which_part(random.random( ), w)
    try:
        return data[c]
    except IndexError:
        return data[-1]

# import random		
# print('weighting_choice', weighting_choice(['a', 'b', 'c', 'd'], [0.536, 0.698, 0.985, 0.1]))
 

# SetUpGeneratorControl()
# 
# RemoveOvercurrentRelay()
# RemoveOverFrequencyGeneratorTripping()
# RemoveUnderFrequencyGeneratorTripping()
# RemoveUnderFrequencyLoadShedding()

# AddOvercurrentRelay()
# ChangeLineRating()
# AddUnderFrequencyLoadShedding()
# AddOverFrequencyGeneratorTripping()
# AddUnderFrequencyGeneratorTripping()

# DisableAVR()
# ChangeSymType('classical')
# ChangeLoadType()
ChangeLoadingLevel(0.65)
ChangeInitialGenerationLevel(1.0)

# GenerateRandomCases(1000)
RunN2Contingencies()
# matching()
