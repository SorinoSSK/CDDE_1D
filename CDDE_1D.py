import pandas as pd
import numpy as np
import pingouin as pg
import pyDOE2 as doe

n = 4
# Read from excel sheet.
BladeStats = pd.read_excel('C:/Users/sweea/.spyder-py3/1D_Data.xlsx')
# print(BladeStats)
# print(BladeStats.loc[(BladeStats['Blade_Scale'] == 3)])
print("---- Factorial Design----")
print("Blade_Scale Relationship:")
Sum_BladeScaleLow = np.sum(BladeStats.loc[(BladeStats['Blade_Scale'] == 1), "Force"].to_numpy())
BS_R_minus = Sum_BladeScaleLow/(4*n)
print("R_-: ", BS_R_minus)
Sum_BladeScaleHigh = np.sum(BladeStats.loc[(BladeStats['Blade_Scale'] == 3), "Force"].to_numpy())
BS_R_plus = Sum_BladeScaleHigh/(4*n)
print("R_+: ", BS_R_plus)
BS_ME = BS_R_plus - BS_R_minus
print("Main Effect, M.E: ", BS_ME)
print("---")
print("No_Of_Blade Relationship:")
Sum_NoOfBladeLow = np.sum(BladeStats.loc[(BladeStats['No_Of_Blade'] == 2), "Force"].to_numpy())
NOB_R_minus = Sum_NoOfBladeLow/(4*n)
print("R_-: ", NOB_R_minus)
Sum_NoOfBladeHigh = np.sum(BladeStats.loc[(BladeStats['No_Of_Blade'] == 4), "Force"].to_numpy())
NOB_R_plus = Sum_NoOfBladeHigh/(4*n)
print("R_+: ", NOB_R_plus)
NOB_ME = NOB_R_plus - NOB_R_minus
print("Main Effect, M.E: ", NOB_ME)
print("---")
print("Blade_Diameter Relationship:")
Sum_BladeDiameterLow = np.sum(BladeStats.loc[(BladeStats['Blade_Diameter'] == 75), "Force"].to_numpy())
BD_R_minus = Sum_BladeDiameterLow/(4*n)
print("R_-: ", BD_R_minus)
Sum_BladeDiameterHigh = np.sum(BladeStats.loc[(BladeStats['Blade_Diameter'] == 100), "Force"].to_numpy())
BD_R_plus = Sum_BladeDiameterHigh/(4*n)
print("R_+: ", BD_R_plus)
BD_ME = BD_R_plus - BD_R_minus
print("Main Effect, M.E: ", BD_ME)






print()
print("---- To Double Check ----")
# Perform multi-factor ANOVA with factor = 2 (ss_type = 2) and round BladeDissectues to 3 decimal point.
# Values returned as panda.dataframe
BladeDissect = BladeStats.anova(dv='Force', between=['Blade_Scale','No_Of_Blade','Blade_Diameter'], ss_type=1).round(3)
print("----Blade_Scale---- ")
print("Degree Of Freedom: ", BladeDissect.iloc[0][2])
print("SS(Blade_Scale)_treatment: ", BladeDissect.iloc[0][1])
print("MS(Blade_Scale): ", BladeDissect.iloc[0][3])
print("F Value: ", BladeDissect.iloc[0][4])
print("P Value: ", BladeDissect.iloc[0][5])
print("Contribution by (Blade_Scale): ", BladeDissect.iloc[0][6])
print()
print("----No_Of_Blade---- ")
print("Degree Of Freedom: ", BladeDissect.iloc[1][2])
print("SS(No_Of_Blade)_treatment: ", BladeDissect.iloc[1][1])
print("MS(No_Of_Blade): ", BladeDissect.iloc[1][3])
print("F Value: ", BladeDissect.iloc[1][4])
print("P Value: ", BladeDissect.iloc[1][5])
print("Contribution by (No_Of_Blade): ", BladeDissect.iloc[1][6])
print()
print("----Blade_Diameter---- ")
print("Degree Of Freedom: ", BladeDissect.iloc[2][2])
print("SS(Blade_Diameter)_treatment: ", BladeDissect.iloc[2][1])
print("MS(Blade_Diameter): ", BladeDissect.iloc[2][3])
print("F Value: ", BladeDissect.iloc[2][4])
print("P Value: ", BladeDissect.iloc[2][5])
print("Contribution by (Blade_Diameter): ", BladeDissect.iloc[2][6])
print()
print("----Information----")
print("SS = Sum of Square")
print("MS = Mean Sum of Square")
print("Contribution by Factor: The larger the value, the more contribution it makes on the force.")
print("- Documentation: https://www.statology.org/partial-eta-squared/#:~:text=Partial%20eta%20squared%20is%20a,other%20variables%20in%20the%20model.")
print(".anova Documentation: https://pingouin-stats.org/generated/pingouin.anova.html")