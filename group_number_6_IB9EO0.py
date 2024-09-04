#!/usr/bin/env python
# coding: utf-8

from pyomo.environ import *
import pandas as pd

# Read the data
df = pd.read_excel("WCG_DataSetV1.xlsx").fillna(0)
week = range(50)
lease_length_index = range(4)
capacity = df.iloc[0,1]
lease_length = [1,4,8,16]
demand = df.iloc[8:58,[5,9,13,17]]
price = df.iloc[8:58,[4,8,12,16]]
returned = pd.concat([df.iloc[8:9,7],df.iloc[8:12,11],df.iloc[8:16,15],df.iloc[8:24,19]], axis=1,ignore_index=True).fillna(0)

# Create concrete model
model = ConcreteModel()
# Define the decision variable
model.x = Var(week,lease_length_index,domain = NonNegativeReals)

# Define the objective function
def Obj_rule(model):
    return sum(7*lease_length[j]*price.iloc[i,j]*model.x[i,j] for i in week for j in lease_length_index)
model.objective = Objective(rule=Obj_rule, sense = maximize)

# Define the demand constraints
def Demand_rule(model,i,j):
    return model.x[i,j] <= demand.iloc[i,j]
model.demand_cons = Constraint(week,lease_length_index, rule=Demand_rule)

# Define the capacity constraints
def Cap_rule(model,i):
    return  sum(model.x[i,j] for j in lease_length_index) + sum(model.x[i-k,j] for j in lease_length_index for k in range(1,4) if j >= 1 and i >= k) + sum(model.x[i-k,j] for j in lease_length_index for k in range(4,8) if j >= 2 and i >= k) + sum(model.x[i-k,j] for j in lease_length_index for k in range(8,16) if j >= 3 and i >= k) + sum(returned.iloc[k,j] for k in range(i,16) for j in lease_length_index if i < 16) <= capacity 
model.cap_cons = Constraint(week, rule=Cap_rule)

# Solve the problem
results = SolverFactory('glpk').solve(model)
results.write()

# Print the results
if 'ok' == str(results.Solver.status):
    print("Total Revenue = ", round(model.objective(),2))
    print("\nVariables Table:")
    var_df = pd.DataFrame({'1-week': model.x[:,0].value, '4-week': model.x[:,1].value, '8-week': model.x[:,2].value, '16-week': model.x[:,3].value})
    var_df.index = range(2,52)
    print(var_df)




