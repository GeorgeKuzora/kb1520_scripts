#open exel file
#check input parametrs
#for each row create a list with needed parametrs
#if there is double-material then set right coeficient

import pandas as pd
import numpy as np


class Exl_upload:
    def __init__(self, file_name):
        self.file_name = file_name
        self.upload_list = []
        self.df_full = None

    def create_df(self):
        with pd.ExcelFile(self.file_name) as xlsx:
            df = pd.read_excel(xlsx, "BOM", usecols=[0,1,2,3,4,5,6,7,8,9,10])
        mult_index = []
        mult_list = []
        for i in range(len(df)):
            try:
                if df.at[i - 1, "Structure Level"] < df.at[i, "Structure Level"]:
                    mult_index.append(df.at[i - 1, "Quantity"])
                elif df.at[i - 1, "Structure Level"] > df.at[i, "Structure Level"]:
                    for i in range(df.at[i - 1, "Structure Level"] - df.at[i, "Structure Level"]):
                        mult_index.pop()
                else:
                    pass
            except ValueError:
                pass
            except KeyError:
                pass
            mult = np.prod(mult_index)
            mult_list.append(mult)   
        multiply = pd.Series(mult_list, name="Multiply")
        self.df_full = pd.concat([df, multiply], axis=1)
        return self.df_full
    


a = Exl_upload("Excel/II78917_2_O.xlsx")
a.create_df()

