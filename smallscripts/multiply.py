'''Короткий скрипт для расчета полного колличества компонентов в составе списка материалов'''

import pandas as pd
import numpy as np

class Exl_upload:
    '''Класс для загрузки и обработки данных. Принимает в качестве атрибута имя файла со списком материалов.'''

    def __init__(self, file_name):
        self.file_name = file_name
        self.upload_list = []
        self.df_full = None

    def create_df(self):
        '''Этот метод создает датафрейм pandas и расчитывает реальное количество компонентов в листе материалов. Реальное количество компонентов расчитывается на основании структуры уровней материалов входящих друг в друга и их количеств. Для расчета использована структура stack. После расчета данные заносятся в Эксель файл'''

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
        self.df_full["True quantity"] = (self.df_full["Quantity"] * self.df_full["Multiply"])
        self.df_full.to_excel("true_" + self.file_name, sheet_name="True_BOM")
        return self.df_full
    
file_name = input("Please enter a file name: ")
a = Exl_upload(file_name)
a.create_df()
