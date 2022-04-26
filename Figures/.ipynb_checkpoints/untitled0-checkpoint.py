# -*- coding: utf-8 -*-
"""
Created on Sun Apr 24 13:52:43 2022

@author: kakdemi
"""

#Importing necessary packages
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

#Defining demand parameters
seasons = ['Winter','Spring','Summer','Fall']
hour_type = ['Peak','Offpeak']
demand_type = ['Average','Maximum']

#Defining the simulation years, generation types and all cases
years = [2023,2030,2040]
gen_types = ['Coal','Natural Gas','Fuel Oil','Nuclear']
All_cases = []
for S in seasons:
    for H in hour_type:
        for D in demand_type:
            Case_name = '{}_{}_{}'.format(S,H,D)
            All_cases.append(Case_name)

# #Reading data for the plots and plotting the generation mix
for year in years:
    
    try:

        Dispatchable=pd.read_excel('../Results/Results_BAU_{}_specific_reserve_margin.xlsx'.format(year),sheet_name='DispatchableGen_MWh',header=0)
#     Nuclear= pd.read_excel('../Results/Results_BAU_{}_specific_reserve_margin.xlsx'.format(year),sheet_name='NuclearGen_MWh',header=0)
#     All_generation_cases= pd.DataFrame(np.zeros((len(All_cases),len(gen_types)),columns=gen_types)

    except FileNotFoundError:
        pass