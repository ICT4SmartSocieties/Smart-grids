from openpyxl import load_workbook
import itertools as it
import matplotlib.pyplot as plt

#import from excel file

pnom_bat = 1000
enom_bat = 1000
soc2 = 50
eta = 0.9
soc_min = 0.1
soc_max = 0.9
p1=0
p2=0


#data collecting
wb = load_workbook(filename='load_gen_profiles.xlsx')
data = wb["Sheet1"]
loads = []
gens = []
      
for row in range(2,data._current_row+1):
    loads.append(float(data["A"+str(row)].value))
    gens.append(float(data["B"+str(row)].value))

#track for plotting
pnetv =[]
p2v   =[]

#smart management
for g,l in it.izip(gens,loads):
    pnet = g - l    #calculation of net power
    pnetv.append(pnet)
    #step1
    if pnet >= pnom_bat: #in the network there is more power and battery is full
        newp1 = pnom_bat #the battery can use all its nominal power
    else:
        newp1= pnet      #at the in, the battery see the net usage
    
    '''
    pnet>0, there is ore energy than t s needed -> CHARGING
    pnet<0, the network ask more energy         -> DISCHARGING
    '''
    #step2
    soc1 = soc2  #the new soc1 is the previuous soc2
#    if the battery is full, and there more enery than is need, the battery does not gie enegy
#    or
#    the battery is empity and there is power demand, the battery cant give enery
#    and soc does not change
    if (soc1 >= soc_max and pnet > 0) or (soc1 < soc_min and pnet < 0):
        p2 = 0
        p2v.append(p2)
        soc1=soc2
    else:
        #step3 - soc different from full or empity
        if  pnet >= 0: #the net power is >0 -> i can charge
            soc2 = soc1 + newp1*eta*15/enom_bat
            p2=0
            p2v.append(p2)
        else: 
            #pnet < 0: #the net power is <0 -> there is power req, i give power
            soc2 = soc1 -(newp1/eta)*(15/enom_bat)
            if abs((newp1/eta)) > pnom_bat:
                p2 = pnom_bat
                p2v.append(p2)
            else :
                p2 = abs((newp1/eta))
                p2v.append(p2)

    #step4
    if soc2 > soc_max: #check if there i fill up the battery
        soc2 = soc_max
        p2 = newp1 

plt.plot(p2v,'r',pnetv,'b')

    
        
        
    
    
        

        