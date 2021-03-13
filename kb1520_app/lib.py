'''Библиотека модулей для программы по подготовке списка материалов для продукта проходящего ремонт'''

import pandas as pd
import numpy as np
import openpyxl

class Data_list:
    '''Класс для загузки данных о расходе материалов в прошлом и существующих продуктах и их свойствах. Требует подготовленных и очищенных данных в формате .xlsx.'''
    
    def __init__(self, basic_file, plant_file, zparts_file, product_file):
        self.basic_file = basic_file
        self.plant_file = plant_file
        self.zparts_file = zparts_file
        self.product_file = product_file
        self.data_df = None
        self.basic_df = None
        self.plant_df = None
        self.zparts_df = None
        self.product_df = None
        self.components_df = None

    def create_basic(self):
        '''Создает датафрейм из листа Эксель'''
        with pd.ExcelFile(self.basic_file) as xlsx:
            self.basic_df = pd.read_excel(xlsx, "basic_query", na_values=["NA"], index_col=0, usecols="A:P")
        return self.basic_df

    def create_plant(self):
        '''Создает датафрейм из листа Эксель'''
        with pd.ExcelFile(self.plant_file) as xlsx:
            self.plant_df = pd.read_excel(xlsx, "plant_query", na_values=["NA"], index_col=0, usecols="A:E")
        return self.plant_df

    def create_zparts(self):
        '''Создает датафрейм из листа Эксель'''
        with pd.ExcelFile(self.zparts_file) as xlsx:
            self.zparts_df = pd.read_excel(xlsx, "zparts_query", na_values=["NA"], usecols="B:Z")
        return self.zparts_df

    def create_product(self):
        '''Создает датафрейм из листа Эксель'''
        with pd.ExcelFile(self.product_file) as xlsx:
            self.product_df = pd.read_excel(xlsx, "product_query", na_values=["NA"], index_col=0, usecols="A:K")
        return self.product_df

    def connect_material(self):
        '''Соединяет два датафрейма и получает датафрейм с данными о используемых компонентах'''
        frames = [self.basic_df, self.plant_df]
        self.components_df = pd.concat(
                                    frames,
                                    axis=1,
                                    join="outer",
                                    ignore_index=False,
                                    )
        return self.components_df

    def connect_repairs(self):
        '''Соединяет два датафрейма и получает датафрейм с данными о проведенных ремонтах и расходе материалов'''
        self.repairs_df = pd.merge(
                                self.zparts_df,
                                self.product_df,
                                how="left",
                                on=None,
                                left_on="Material",
                                right_on="Service material",
                                left_index=False,
                                right_index=True,
                                sort=False,
                                suffixes=("_x", "_y"),
                                copy=True,
                                indicator=False,
                                validate=None,
                                )
        return self.repairs_df


class User_data:
    '''Класс для загрузки и обработки данных загружаеммых пользователем, а именно ВОМ материала. Принимает в качестве атрибута имя файла со списком материалов. Требует очистки данных пользователя'''

    def __init__(self, file_name):
        self.file_name = file_name
        self.upload_list = []
        self.user_df = None

    def create_df(self):
        '''Этот метод создает датафрейм pandas и расчитывает реальное количество компонентов в листе материалов. Реальное количество компонентов расчитывается на основании структуры уровней материалов входящих друг в друга и их количеств. Для расчета использована структура stack.''' 

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
        self.user_df = pd.concat([df, multiply], axis=1)
        self.user_df["True quantity"] = (self.user_df["Quantity"] * self.user_df["Multiply"])
        return self.user_df

    
if __name__ == '__main__':

#    a = Data_list(
#            "/home/georgiy/Documents/Code/kb1520_app/Excel/data/query/basic_query.xlsx",
#            "/home/georgiy/Documents/Code/kb1520_app/Excel/data/query/plant_query.xlsx",
#            "/home/georgiy/Documents/Code/kb1520_app/Excel/data/query/zparts_query.xlsx",
#            "/home/georgiy/Documents/Code/kb1520_app/Excel/data/query/product_query.xlsx",
#            )
#    a.create_basic()
#    a.create_plant()
#    a.create_zparts()
#    a.create_product()
#
#    print(a.zparts_df)
#    print(a.product_df)
#    a.connect_repairs()
#
#    print(a.repairs_df)

    b = User_data("/home/georgiy/Documents/Code/kb1520_app/Excel/II78917_2_O.xlsx")
    b.create_df()
    print(b.user_df)
