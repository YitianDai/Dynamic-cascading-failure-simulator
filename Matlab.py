def Dynamic_Load_Shedding_Calculator_IEEE39(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
    a = list(map(int,faulty_line))
    b = list(map(int,faulty_node_overfreq))
    c = list(map(int,faulty_node_underfreq))
    d = list(map(int,target_load))
    e = amount_of_shedding
    LS = eng.Dynamic_Load_Shedding_Calculator1_IEEE39(matlab.int32(a),matlab.int32(b),matlab.int32(c),matlab.int32(d),matlab.double(e),nargout = 5)
    return a, b, c, LS[0], LS[1], LS[2], LS[3], LS[4]



# print(Dynamic_Load_Shedding_Calculator(['12', '1', '10', '9', '6', '16'], ['32', '31'], [3, 4, 7, 8, 12, 15, 16, 18, 20, 21, 23, 24, 25, 26, 27, 28, 29, 31, 39], [64.4, 50.0, 11.690000000000001, 26.1, 2.0625, 88.0, 90.47500000000001, 43.45, 172.70000000000002, 75.35000000000001, 68.0625, 84.86500000000001, 44.800000000000004, 27.8, 56.2, 41.2, 56.7, 2.53, 110.4]))



def Dynamic_Load_Shedding_Calculator_ACTIVSg200(loading, faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
 
    a = list(map(int,faulty_line))
    b = list(map(int,faulty_node_overfreq))
    c = list(map(int,faulty_node_underfreq))
    d = list(map(int,target_load))
    e = amount_of_shedding
    LS = eng.Dynamic_Load_Shedding_Calculator_ACTIVSg200(matlab.double([loading]), matlab.int32(a),matlab.int32(b),matlab.int32(c),matlab.int32(d),matlab.double(e),nargout = 5)
    return a, b, c, LS[0], LS[1], LS[2], LS[3], LS[4]


def Dynamic_Load_Shedding_Calculator_ACTIVSg2000(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
    a = list(map(int,faulty_line))
    b = list(map(int,faulty_node_overfreq))
    c = list(map(int,faulty_node_underfreq))
    d = list(map(int,target_load))
    e = amount_of_shedding
    LS = eng.Dynamic_Load_Shedding_Calculator_ACTIVSg2000(matlab.int32(a),matlab.int32(b),matlab.int32(c),matlab.int32(d),matlab.double(e),nargout = 5)
    return a, b, c, LS[0], LS[1], LS[2], LS[3], LS[4]



def Dynamic_form_iteration(a1=[], a2=[], a3=[], a4=[], a5=[], a6=[], a7=[], a8=[], a9=[], a10=[]) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
    a1 = list(map(int,a1))
    a2 = list(map(int,a2))
    a3 = list(map(int,a3))
    a4 = list(map(int,a4))
    a5 = list(map(int,a5))
    a6 = list(map(int,a6))
    a7 = list(map(int,a7))
    a8 = list(map(int,a8))
    a9 = list(map(int,a9))
    a10 = list(map(int,a10))
    L = eng.dynamic_form_iteration(matlab.int32(a1),matlab.int32(a2),matlab.int32(a3),matlab.int32(a4),matlab.int32(a5),matlab.int32(a6),matlab.int32(a7),matlab.int32(a8),matlab.int32(a9),matlab.int32(a10),nargout = 1)
    return L

# print(Dynamic_form_iteration(['3191', '367', '18'], ['3193'], ['3166', '3056', '3089'], ['3014', '3122']))