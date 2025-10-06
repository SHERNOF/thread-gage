import xlwings as xw

# --------------- Interface ----------------------------------------------- #

# from tkinter import *
# from tkinter import ttk
# root = Tk()
# frm = ttk.Frame(root, padding=10)
# frm.grid()
# ttk.Label(frm, text="Hello World!").grid(column=0, row=0)
# ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=0)
# root.mainloop()

# --------------- Interface ----------------------------------------------- #

# wb = xw.books.open(r"/Volumes/DDrive/PhytonZTM/Tex At Site Thread Gage Worksheet.xlsm")
wb = xw.books.open(r'C:\Tex Onsite\Tex_reff\shernof\thread-gage\Tex At Site Thread Gage Worksheet.xlsm')
#wb = xw.books.open(r'C:\Tex Onsite\Tex_reff\M25X1.5-6H.xlsm')

td = wb.sheets("Test Data")
vr = wb.sheets("Variables")
import math


def float_range(start, stop, step):  
  while start < stop:
      yield round(start, 1)
      start += step

def float_range_imp(start, stop, step):  
  while start < stop:
      yield round(start, 0.1)
      start += step

"""
From Table 4.6 of AS1014 Deviations for Go Screw Plug Gauges
"""
      
deviation_table = [
        [range(24, 51),{'du':3,'dl':-3}, {'worn':-8},{'mu':6},{'ml':-6}],
        [range(51, 81),{'du':6},{'dl':-1}, {'worn':-7},{'mu':9},{'ml':-5}],
        [range(81, 126),{'du':11},{'dl':2}, {'worn':-6},{'mu':15},{'ml':-3}],
        [range(126, 201),{'du':18},{'dl':6}, {'worn':-5},{'mu':23},{'ml':1}],
        [range(201, 316),{'du':23},{'dl':9}, {'worn':-5},{'mu':30},{'ml':2}],
        [range(316, 501),{'du':33},{'dl':15}, {'worn':-3},{'mu':42},{'ml':6}],
        [range(501, 671),{'du':43},{'dl':21}, {'worn':-1},{'mu':54},{'ml':10}],
        ]
        
        
"""
From Table 4.7 of AS1014 Deviations for No-Go Screw Plug Gauges
"""
deviation_table_nogo = [
        [range(24, 51),{'du':6},{'dl':0}, {'worn':-3},{'mu':9},{'ml':-3}],
        [range(51, 81),{'du':7},{'dl':0}, {'worn':-4},{'mu':10},{'ml':-4}],
        [range(81, 126),{'du':9},{'dl':0}, {'worn':-5},{'mu':13},{'ml':-5}],
        [range(126, 201),{'du':11},{'dl':0}, {'worn':-6},{'mu':16},{'ml':-6}],
        [range(201, 316),{'du':14},{'dl':0}, {'worn':-8},{'mu':21},{'ml':-7}],
        [range(316, 501),{'du':18},{'dl':0}, {'worn':-10},{'mu':27},{'ml':-9}],
        [range(501, 671),{'du':22},{'dl':0}, {'worn':-12},{'mu':33},{'ml':-11}],
        ]

"""
From Table 4.2 of AS1721 Pitch Diamaeter Tolerance for Internal Threads
"""
tolerance_table = [
        [float_range(0.99,1.41,0.01),(0.2),{'4':40,'5':'','6':'','7':'','8':''}],
        [float_range(0.99,1.41,0.01),(0.25),{'4':45,'5':56,'6':'','7':'','8':''}],
        [float_range(0.99,1.41,0.01),(0.3),{'4':48,'5':60,'6':75,'7':'','8':''}],
        
        [float_range(1.41,2.81,0.1),(0.2),{'4':42,'5':'','6':'','7':'','8':''}],
        [float_range(1.41,2.81,0.1),(0.25),{'4':48,'5':60,'6':'','7':'','8':''}],
        [float_range(1.41,2.81,0.1),(0.35),{'4':53,'5':67,'6':85,'7':'','8':''}],
        [float_range(1.41,2.81,0.1),(0.4),{'4':56,'5':71,'6':90,'7':'','8':''}],
        [float_range(1.41,2.81,0.1),(0.45),{'4':60,'5':75,'6':95,'7':'','8':''}],
        
        [float_range(2.81,5.61,0.1),(0.35),{'4':56,'5':71,'6':90,'7':'','8':''}],
        [float_range(2.81,5.61,0.1),(0.5),{'4':63,'5':80,'6':100,'7':125,'8':''}],
        [float_range(2.81,5.61,0.1),(0.6),{'4':71,'5':90,'6':112,'7':140,'8':''}],
        [float_range(2.81,5.61,0.1),(0.7),{'4':75,'5':95,'6':118,'7':150,'8':''}],
        [float_range(2.81,5.61,0.1),(0.75),{'4':75,'5':95,'6':118,'7':150,'8':''}],
        [float_range(2.81,5.61,0.1),(0.8),{'4':80,'5':100,'6':125,'7':160,'8':200}],
        
        [float_range(5.61,11.21,0.1),(0.75),{'4':85,'5':106,'6':132,'7':170,'8':''}],
        [float_range(5.61,11.21,0.1),(1),{'4':95,'5':118,'6':150,'7':190,'8':236}],
        [float_range(5.61,11.21,0.1),(1.25),{'4':100,'5':125,'6':160,'7':200,'8':250}],
        [float_range(5.61,11.21,0.1),(1.5),{'4':112,'5':140,'6':180,'7':224,'8':280}],
        
        [float_range(11.21,22.41,0.1),(1),{'4':100,'5':125,'6':160,'7':200,'8':250}],
        [float_range(11.21,22.41,0.1),(1.25),{'4':112,'5':140,'6':180,'7':224,'8':280}],
        [float_range(11.21,22.41,0.1),(1.5),{'4':118,'5':150,'6':190,'7':236,'8':300}],
        [float_range(11.21,22.41,0.1),(1.75),{'4':125,'5':160,'6':200,'7':250,'8':315}],
        [float_range(11.21,22.41,0.1),(2),{'4':132,'5':170,'6':212,'7':265,'8':335}],
        [float_range(11.21,22.41,0.1),(2.5),{'4':140,'5':180,'6':224,'7':280,'8':355}],
        
        [float_range(22.41,46,0.1),(1),{'4':106,'5':132,'6':170,'7':212,'8':''}],
        [float_range(22.41,46,0.1),(1.5),{'4':125,'5':160,'6':200,'7':250,'8':315}],
        [float_range(22.41,46,0.1),(2),{'4':140,'5':180,'6':224,'7':280,'8':355}],
        [float_range(22.41,46,0.1),(3),{'4':170,'5':212,'6':265,'7':335,'8':425}],
        [float_range(22.41,46,0.1),(3.5),{'4':180,'5':224,'6':280,'7':355,'8':460}],
        [float_range(22.41,46,0.1),(4),{'4':190,'5':236,'6':300,'7':375,'8':475}],
        [float_range(22.41,46,0.1),(4.5),{'4':200,'5':250,'6':315,'7':400,'8':500}],
        
        [range(46,91),(1.5),{'4':132,'5':170,'6':212,'7':265,'8':335}],
        [range(46,91),(2),{'4':150,'5':191,'6':236,'7':300,'8':375}],
        [range(46,91),(3),{'4':180,'5':224,'6':280,'7':355,'8':460}],
        [range(46,91),(4),{'4':200,'5':250,'6':315,'7':400,'8':500}],
        [range(46,91),(5),{'4':212,'5':265,'6':335,'7':425,'8':530}],
        [range(46,91),(5.5),{'4':224,'5':280,'6':355,'7':460,'8':560}],
        [range(46,91),(6),{'4':236,'5':300,'6':375,'7':475,'8':600}],
        
        [range(91,181),(2),{'4':160,'5':200,'6':250,'7':315,'8':400}],
        [range(91,181),(3),{'4':190,'5':236,'6':300,'7':375,'8':475}],
        [range(91,181),(4),{'4':212,'5':265,'6':335,'7':425,'8':530}],                                                   
        [range(91,181),(6),{'4':250,'5':315,'6':400,'7':500,'8':630}],
        
        [range(181,356),(3),{'4':212,'5':265,'6':335,'7':425,'8':530}],
        [range(181,356),(4),{'4':236,'5':300,'6':375,'7':475,'8':600}],
        [range(181,356),(6),{'4':265,'5':355,'6':425,'7':530,'8':670}],                                           
        ]

"""
from table 6 of ASME B1.2-1983
X Gage Tolerance for Thread Gages
"""

imperial_tolerance_major_minorD = [
    [float_range(0.0001,4,0.0001),{80:0.0003,72:0.0003,64:0.0004,56:0.0004,48:0.0004,44:0.0004,40:0.0004,36:0.0004,32:0.0005,28:0.0005,24:0.0005,20:0.0005,18:0.0005,16:0.0006,14:0.0006,13:0.0006,12:0.0006,11.5:0.0006,11:0.0006,10:0.0006,9:0.0007,8:0.0007,7:0.0007,6:0.0008,5:0.0008,4.5:0.0008,4:0.0009}],
    [float_range(4,100,0.0001),{80:0.0000,72:0.0000,64:0.0000,56:0.0000,48:0.0000,44:0.0000,40:0.0000,36:0.0000,32:0.0007,28:0.0007,24:0.0007,20:0.0007,18:0.0007,16:0.0009,14:0.0009,13:0.0009,12:0.0009,11.5:0.0009,11:0.0009,10:0.0009,9:0.0011,8:0.0011,7:0.0011,6:0.0013,5:0.0013,4.5:0.0013,4:0.0015}],
]

"""
from table 6 of ASME B1.2-1983
X Gage Tolerance for Thread Gages
"""
imperial_tolerance_pitch_diameter = [
    [float_range(0.0000,1.5001,0.0001),{80:0.0002,72:0.0002,64:0.0002,56:0.0002,48:0.0002,44:0.0002,40:0.0002,36:0.0002,32:0.0003,28:0.0003,24:0.0003,20:0.0003,18:0.0003,16:0.0003,14:0.0003,13:0.0003,12:0.0003,11.5:0.0003,11:0.0003,10:0.0003,9:0.0003,8:0.0004,7:0.0004,6:0.0004,5:0.0000,4.5:0.0000,4:0.0000}],
    [float_range(1.5,4.001,0.0001),{80:0.0000,72:0.0000,64:0.0000,56:0.0003,48:0.0003,44:0.0003,40:0.0003,36:0.0003,32:0.0004,28:0.0004,27:0.0004,24:0.0004,20:0.0004,18:0.0004,16:0.0004,14:0.0004,13:0.0004,12:0.0004,11.5:0.0004,11:0.0004,10:0.0004,9:0.0004,8:0.0005,7:0.0005,6:0.0005,5:0.0005,4.5:0.0005,4:0.0005}],
    [float_range(4,8,0.0001),{80:0.0000,72:0.0000,64:0.0000,56:0.0000,48:0.0000,44:0.0000,40:0.0000,36:0.0000,32:0.0005,28:0.0005,27:0.0005,24:0.0005,20:0.0005,18:0.0005,16:0.0006,14:0.0006,13:0.0006,12:0.0006,11.5:0.0006,11:0.0006,10:0.0006,9:0.0006,8:0.0006,7:0.0006,6:0.0006,5:0.0006,4.5:0.0006,4:0.0006}],
]

"""
from table 4.4 of ASME1721 - Minor Diamater tolerancs for Internal Threads
2nd element is from Table 3.1 of AS1014 under 0.75H column
3rd element is from Table 4.1 of AS1721 - Fundamental Deviations for G
"""

pitch_34p_fdforG_minord = [
        [{'pitch':0.2},(0.13),(17),{'4':38,'5':'','6':'','7':'','8':''}],
        [{'pitch':0.25},(0.162),(18),{'4':45,'5':56,'6':'','7':'','8':''}],
        [{'pitch':0.3},(0.195),(18),{'4':53,'5':67,'6':85,'7':'','8':''}],
        [{'pitch':0.35},(0.227),(19),{'4':63,'5':80,'6':100,'7':'','8':''}],
        [{'pitch':0.4},(0.26),(19),{'4':71,'5':90,'6':112,'7':'','8':''}],
        [{'pitch':0.45},(0.292),(20),{'4':80,'5':100,'6':125,'7':'','8':''}],
        [{'pitch':0.5},(0.325),(20),{'4':90,'5':112,'6':140,'7':180,'8':''}],
        [{'pitch':0.6},(0.39),(21),{'4':100,'5':125,'6':160,'7':200,'8':''}],
        [{'pitch':0.7},(0.455),(22),{'4':112,'5':140,'6':180,'7':224,'8':''}],
        [{'pitch':0.75},(0.487),(22),{'4':118,'5':150,'6':190,'7':236,'8':''}],
        [{'pitch':0.8},(0.52),(24),{'4':125,'5':160,'6':200,'7':250,'8':315}],
        [{'pitch':1},(0.65),(26),{'4':150,'5':190,'6':236,'7':300,'8':375}],
        [{'pitch':1.25},(0.812),(28),{'4':170,'5':212,'6':265,'7':335,'8':425}],
        [{'pitch':1.5},(0.974),(32),{'4':190,'5':236,'6':300,'7':375,'8':475}],
        [{'pitch':1.75},(1.137),(34),{'4':212,'5':265,'6':335,'7':425,'8':530}],
        [{'pitch':2},(1.3),(38),{'4':236,'5':300,'6':375,'7':475,'8':600}],
        [{'pitch':2.5},(1.624),(42),{'4':280,'5':355,'6':450,'7':560,'8':710}],
        [{'pitch':3},(1.95),(48),{'4':315,'5':400,'6':500,'7':630,'8':800}],
        [{'pitch':3.5},(2.273),(53),{'4':355,'5':450,'6':560,'7':710,'8':900}],
        [{'pitch':4},(2.598),(60),{'4':375,'5':475,'6':600,'7':750,'8':950}],
        [{'pitch':4.5},(2.923),(63),{'4':425,'5':530,'6':670,'7':850,'8':1060}],
        [{'pitch':5},(3.25),(71),{'4':450,'5':560,'6':710,'7':900,'8':1120}],
        [{'pitch':5.5},(3.572),(75),{'4':475,'5':600,'6':750,'7':950,'8':'1180'}],
        [{'pitch':6},(3.897),(80),{'4':500,'5':630,'6':800,'7':1000,'8':1250}],        
    ]


# 5.2 of B1.1; this is to define the factor to be use to compute the LE; whether 9P or 1D
table_2_list = [12,16,20,28,32]


def thread_gage(): 
    unitCheck = td.range('D9').value
    if unitCheck == 'mm':
        model = td.range('D11').value
        
        x = model.index('X')
        dsh = model.index('-')
        dia = int(model[1:x])
        pd = float(model[x + 1:dsh])
        tol = str(model[dsh + 1])
        gauge = model[len(model)-1:len(model)]
        vr.range('Q3').value = ''
        vr.range('Q5').value = dia
        vr.range('Q6').value = pd
        vr.range('Q7').value = tol
        vr.range('Q8').value = gauge
        vr.range('R4').value = 'Tolerance'
        vr.range('S4').value = 'Index'
        vr.range('Q4').value = model
        vr.range('P6').value = ""

        for i in tolerance_table:
            
            if dia in i[0]:
                # if pd in i[1].values(): #acquire the value of the pitch D
                if pd in i:
                    if tol in i[2]:
                        vr.range('R5').value = i[2].get(tol) #getting the value using the key
                        tolerance = i[2].get(tol)

                        # getting the deviations from go table
                        #test = [row[0] for row in deviation_table] 

                        # deviations for go
                        for go, rng in  enumerate(deviation_table):
                            if tolerance in rng[0]:
                                
                                vr.range('V12').value = (deviation_table[go][2]["dl"])/1000
                                vr.range('V13').value = (deviation_table[go][1]["du"])/1000
                                vr.range('X12').value = (deviation_table[go][5]["ml"])/1000
                                vr.range('X13').value = (deviation_table[go][4]["mu"])/1000
                                vr.range('Z12').value = (deviation_table[go][3]["worn"])/1000
                        # deviations for no-go
                        for nogo, rng in enumerate(deviation_table_nogo):
                            if tolerance in rng[0]:
                                vr.range('W12').value = (deviation_table_nogo[nogo][2]["dl"])/1000
                                vr.range('W13').value = (deviation_table_nogo[nogo][1]["du"])/1000
                                vr.range('Y12').value = (deviation_table_nogo[nogo][5]["ml"])/1000
                                vr.range('Y13').value = (deviation_table_nogo[nogo][4]["mu"])/1000
                                vr.range('AA12').value = (deviation_table_nogo[nogo][3]["worn"])/1000
                        

                        for i, rng in enumerate(deviation_table):
                            if tolerance in rng[0]:
                                vr.range('S5').value = i # deviation index
                                
        # minor diameter
        for item in pitch_34p_fdforG_minord:
            if pd in item[0].values():
                if tol in item[3]:
                    vr.range('AA4').value = item[3].get(tol) #acquire the key for the minor diameter
        
        #for G
        for item in pitch_34p_fdforG_minord:
            if gauge == 'G':
                if pd in item[0].values():
                    vr.range('R8') .value = item[2]/1000
            else:
                vr.range('R8') .value = ''

    else:
        
        print('Thank you Lord')
        # vr.range('P4').value = 'Metric Equivalent'
        model = td.range('D11').value
        x = model.index('X')
        space1 = model.index(' ')
        space2 = model.find(' ',2)
        # sl = model.index('/')
        # print(sl)

        try: 
            model.index('/') # this is to test if the diameter > 1"

        except ValueError:
            dia_converted = round(int(model[0:x])  * 25.4)
            dia = int(model[0:x])
        else:
            if model.index('/') == 1:  #1/4X28 UNF-2B
                dia_converted = round(int(model[0]) / int(model[2:x]) * 25.4)
                dia = int(model[0]) / int(model[2:x])
                vr.range('P5').value = dia

            elif model.index('/') == -1: #1X8 UNC-2B
                dia_converted = round(int(model[0]) * 25.4)
                dia = int(model[0]) 
                vr.range('P5').value = dia
         
            elif model.index('/') == 2: #15/16X20 UNEF-2B
                dia_converted = round(int(model[0:1]) / int(model[3:x]) * 25.4)
                dia = int(model[0:1]) / int(model[3:x]) 
                vr.range('P5').value = dia

            elif model.index('/') == 3: #1 1/4X7 UNF-2B
                dia_converted = round( int(model[0]) + int(model[2:model.index('/')]) / int(model[model.index('/') + 1:x]) + int(model[0]) * 25.4)
                dia = (int(model[0]) + int(model[2:model.index('/')]) / int(model[model.index('/') + 1:x]))
                vr.range('P5').value = dia

        
    # Getting the pitch diameter
        space2
        pd = round(float(model[x + 1:space2]) if space2 else float(model[x + 1:space1]))
        for i in table_2_list:
            if pd in table_2_list:
                vr.range('R4').value = "Yes"
            else:
                vr.range('R4').value = "No"

    # converting TPI to mm
        pd_converted = 1 /float(model[x + 1:space2]) * 25.4 if space2 else float(model[x + 1:space1] * 25.4) 
        pdtm = 1 / float(model[x + 1:space2]) * 25.4 if space2 else 1 / float(model[x + 1:space1])
        test = [row[1] for row in tolerance_table] 

        """
        the pdtm column was separated and get the minimum difference to get the nearest equivalent
        of the pitch diameter. 
         """
        pd_list = []
        for i in test: 
            """
            this is the loop of all the pitch diameter and put inside the test[] to be able to get the difference of each pitch vs the actual pitch of the IUT
            the minimum difference is the appropriate match of the converted TPI to mm
            """
            pd_diff = abs(pdtm - i) #difference between each available pitch versus the IUT pitch
            pd_list.append(pd_diff)
            pd_index =pd_list.index(min(pd_list))
            available_pitch = test[pd_index]
        pdtm > available_pitch
        pd_converted = abs(min(pd_list)-pdtm) if pdtm > available_pitch else abs(min(pd_list)+pdtm)

        # getting the tolerance
        dsh = model.index('-')
        toltm = model[dsh + 1:]
        if toltm == '1B':
            tol = '7'
        elif toltm == '2B':
            tol = '6'
        else:
            toltm = '3B'
            tol = '5'
        
        """
        Getting the constants for the Pitch Diameter
        """
        
        for i in imperial_tolerance_pitch_diameter:
            if int(dia) in i[0]:
                """
                if pd in i[1].values(): #acquire the value of the pitch D
                This was changed because the tolerance_table[1] was changed from a dict to list so that the data can be iterate easily
                """
                if pd in i[1]:
                    vr.range('R5').value = i[1].get(pd) #getting the value using the key
                    print(i[1].get(pd))
        # minor and major diameter tolerance
        for i in imperial_tolerance_major_minorD:
            if int(dia) in i[0]:
                if pd in i[1]:
                    vr.range('S5').value = i[1].get(pd) #acquire the key for the minor diameter
                    print(i[1].get(pd))
    
        converted_model = (f'M{dia_converted}X{pd_converted}-{tol}H')
        vr.range('P4').value = model
        # vr.range('P5').value = dianr #non rounded diameter
        vr.range('P5').value = dia #non rounded diameter
        vr.range('P6').value = pd #non rounded pitch diameter
        vr.range('P7').value = toltm #non rounded pitch diameter        
        vr.range('Q4').value = converted_model
        vr.range('Q5').value = dia_converted
        vr.range('Q6').value = pd_converted
        vr.range('Q7').value = tol
        gauge = converted_model[len(converted_model)-1:len(converted_model)]
        vr.range('Q8').value = gauge
        vr.range('R4').value = 'Pitch D Tolerance'
        vr.range('S4').value = 'Diameter Tolerance'


thread_gage()

print(xw.__version__)
    
    


