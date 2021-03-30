'''Библиотека модулей для программы по подготовке списка материалов для продукта
проходящего ремонт'''

import pandas as pd
import numpy as np
import openpyxl

class Data_list:
    '''Класс для загузки данных о расходе материалов в прошлом и существующих
    компонентах и их свойствах. Требует подготовленных и очищенных данных в формате .xlsx.'''

    def __init__(self, basic_file):
        '''Атрибут имя файла .xlsx где храняться данные разделенные по вкладкам,
        сформированныи при промощи Powerquery'''
        self.basic_file = basic_file

    def create_materials(self):
        '''Создает датафрейм из листа Эксель в котором содержатся данные о
        существующих материалах и их свойствах.'''
        with pd.ExcelFile(self.basic_file) as xlsx:
            self.materials_df = pd.read_excel(xlsx, "material_data", na_values=["NA"], index_col=0, usecols="A:F")
        return self.materials_df

    def create_zparts(self):
        '''Создает датафрейм из листа Эксель в котором содержаться данные о
        произведенных ремонтах и расходе материалов в этих ремонтах.'''
        with pd.ExcelFile(self.basic_file) as xlsx:
            self.zparts_df = pd.read_excel(xlsx, "zparts_combined", na_values=["NA"], usecols="A:Z")
        return self.zparts_df


class User_data:
    '''Класс для загрузки и обработки данных загружаеммых пользователем,
    а именно ВОМ материала и данные о продукте занесенные в пользовательский файл.
    Требует очистки данных пользователя'''

    def __init__(self, file_name, material_df):
        '''Атрибут имя файла .xlsx где храняться данные разделенные по вкладкам,
        сформированныи при промощи Powerquery'''
        self.file_name = file_name
        self.material_df = material_df
        self.coef_array = []
        self.key_id = []
        self.date_id = []
        self.sum_id = []

    def create_df(self):
        '''Этот метод создает датафрейм для материалов входящих в БОМ для которого
        производится расчет потребностей, и расчитывает реальное количество
        компонентов в листе материалов.'''

        with pd.ExcelFile(self.file_name) as xlsx:
            df = pd.read_excel(xlsx, "Default", usecols="A:J")
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

    def create_prdinfo(self):
        '''Создает датафрейм данных о ремонтируемом продукте из листа Эксель,
        загружаемого пользователем. Данные о продукте вносятся пользователем во
        вкладку info'''
        with pd.ExcelFile(self.file_name) as xlsx:
            self.prdinfo_df = pd.read_excel(xlsx, "info", na_values=["NA"], usecols="A:C")
        return self.prdinfo_df

    def create_coef_array(self, coef, key, date, sum_qb):
        '''Добавляет расчитанный для объекта компонента, входящего в лист с
        пользовательскими данными, через который идет итерация, в лист для
        последующего добвавления к данным'''
        self.coef_array.append(coef)
        self.key_id.append(key)
        self.date_id.append(date)
        self.sum_id.append(sum_qb)

    def coef_arr_concat(self):
        '''Присоединяет лист к добавленными коэффициентами к основному дата фрейму
        пользовательских данных'''
        coef_arr_s = pd.Series(self.coef_array, name="Coef")
        key_id_s = pd.Series(self.key_id, name="Key_id")
        date_id_s = pd.Series(self.date_id, name="Date_id")
        sum_id_s = pd.Series(self.sum_id, name="Sum_id")
        self.user_df = pd.concat([self.user_df, coef_arr_s], axis=1)
        self.user_df = pd.concat([self.user_df, key_id_s], axis=1)
        self.user_df = pd.concat([self.user_df, date_id_s], axis=1)
        self.user_df = pd.concat([self.user_df, sum_id_s], axis=1)
        return self.user_df

    def material_data_merge(self):
        material_s = self.material_df["Material supply chain status"]
        print(self.material_df.head())
        self.user_df = pd.merge(
                                self.user_df,
                                material_s,
                                how="left",
                                on=None,
                                left_on="Number",
                                right_on="Material",
                                left_index=False,
                                right_index=True,
                                sort=False,
                                suffixes=("_x", "_y"),
                                copy=True,
                                indicator=False,
                                validate=None,
                                )

    def print_df(self):
        '''Записывает пользовательский датафрейм в новый эксель файл'''
        self.user_df.to_excel("true_" + self.file_name, sheet_name="BOM", index=False)


class Material_obj:
    '''Общий класс для создаваеммых объектов компонентов и продуктов, так как
    они разделяют много схожих параметров, мы воспользуемся принципом наследственности.'''

    def __init__(self, index, user_db_name, data_db_name):
        '''Атрибуты:
           Индекс строки в исходных данных для продукта это
           будет 0, так как в данных одна строчка, для компонента это будет
           индекс строки в датафрейме пользовательских данных через который идет
           итерация.
           Имя данных пользователя: для продукта это имя датафрейма prdinfo_df,
           для компонента это имя датафрейма user_df.
           Имя данных о материале: это имя данных датафрейма materials_df.'''
        self.index = index
        self.user_db_name = user_db_name
        self.data_db_name = data_db_name

    def prepare_objdata(self):
        '''Метод создает новый объект в формате pd.df, из строки заданных данных
        пользователя (либо одна строка для продукта, либо строка через которую
        проходит итерация для компонента'''
        self.objdata_df = self.user_db_name.iloc[[self.index],:]
        return self.objdata_df

    def create_obj(self):
        '''Метод создает новый датафрейм на основании данных пользователя(для
        конкретной строки и общих данных о всех материалах.'''
        self.obj_df = pd.merge(
                                self.objdata_df,
                                self.data_db_name,
                                how="left",
                                on=None,
                                left_on="Number",
                                right_on="Material",
                                left_index=False,
                                right_index=True,
                                sort=False,
                                suffixes=("_x", "_y"),
                                copy=True,
                                indicator=False,
                                validate=None,
                                )
        return self.obj_df


class Component_obj(Material_obj):
    '''Класс дочерний к классу материала, нужен для создания конкретного объекта
    компонента'''

    def rename_obj(self):
        '''Метод переименовывает необходимые столбцы в датафрейме компонента,
        для дальнейшего удобства'''
        self.objr_df = self.obj_df.rename({"Number": "Component",
                            "Material material group": "Component material group",
                            "Material COC": "Component COC",
                            "Material shelf life": "Component shelf life",
                            "Material description": "Component description"}, axis="columns")
        self.objri_df = self.objr_df.reset_index(drop=True)
        return self.objri_df

class Product_obj(Material_obj):
    '''Класс дочерний к классу материала, нужен для создания конкретного объекта
    продукта'''

    def rename_obj(self):
        '''Метод переименовывает необходимые столбцы в датафрейме продукта,
        для дальнейшего удобства'''
        self.objr_df = self.obj_df.rename({"Number": "Product",
                            "Material material group": "Product material group",
                            "Material COC": "Product COC",
                            "Material shelf life": "Product shelf life",
                            "Material description": "Product description"}, axis="columns")
        return self.objr_df


class Complete_df:
    '''Класс для окончательного расчета данных о коэффициенте для каждого из
    компонентов в пользовательских данных, обладает атрибутом класса - словарем
    с данными о названиях столбцов через которые будет идти расширение выборки'''
    SPEC_DICT = {1: ("Component", "Service material"),
                 2: ("Component", "Product"),
                 3: ("Component", "Product description"),
                 4: ("Component", "Product material group"),
                 5: ("Component", "Product COC"),
                 6: ("Component", None),
                 7: ("Component description", "Service material"),
                 8: ("Component description", "Product"),
                 9: ("Component description", "Product description"),
                 10: ("Component description", "Product material group"),
                 11: ("Component description", "Product COC"),
                 12: ("Component description", None),
                 13: ("Component material group", "Service material"),
                 14: ("Component material group", "Product"),
                 15: ("Component material group", "Product description"),
                 16: ("Component marerial group", "Product material group"),
                 17: ("Component material group", "Product COC"),
                 18: ("Component COC", "Service material"),
                 19: ("Component COC", "Product"),
                 20: ("Component COC", "Product description"),
                 21: ("Component COC", "Product material group"),
                 22: ("Component COC", "Product COC")
                }
    DATE_ARR = (2021, 2020, 2019, 2018, 2017, 2016, 2015)

    def __init__(self, zparts_df, component_obj, product_obj):
        '''Атрибуты:
            Данные о ремонтах: данные из датафрейма zparts_df,
            Объект компонента: объект компонента для каждого из материалов в
            пользовательских данных,
            Объект продукта: объект продукта для которого производиться ремонт'''
        self.zparts_df = zparts_df
        self.component_obj = component_obj
        self.product_obj = product_obj
        #self.key_id = None

    def join_obj(self):
        '''Метод соединяет объеты компонента и продукта в один датафрейм,
        для того чтобы было проще создавать выборку из данных о ремонтах'''
        frames = [self.component_obj, self.product_obj]
        self.joined_obj = pd.concat([self.component_obj, self.product_obj], axis=1)
        return self.joined_obj

    def create_query(self):
        '''Метод итерирует через данные словаря для расширения выборки. Для
        каждой из итераций создает выборку с данными о ремонтах продуктов и
        колличестве потребленного компонента в данных ремонтах. При наличии более
        20 ремонтов в выборке, возвращает созданную на данном шаге выборку.'''
        for key, value in self.SPEC_DICT.items():
            component_id = None
            product_id = None
            try:
                component_id = self.joined_obj.at[0, value[0]]
            except (IndexError, KeyError):
                pass
            try:
                product_id = self.joined_obj.at[0, value[1]]
            except (IndexError, KeyError):
                pass
            for date in self.DATE_ARR:
                date_df = self.zparts_df[(self.zparts_df["Date"] >= date)]
                if component_id is None and product_id is not None:
                    self.current_query = date_df[(date_df[value[1]] == product_id)]
                elif product_id is None and component_id is not None:
                    self.current_query = date_df[(date_df[value[0]] == component_id)]
                elif product_id is None and component_id is None:
                    self.current_query = date_df
                else:
                    self.current_query = date_df[(date_df[value[0]] == component_id)
                                                        & (date_df[value[1]] == product_id)]
                sum_qb = self.current_query["Quantity Balance"].sum()
                if sum_qb >= 20:
                    self.calculate_date(date)
                    break
            if sum_qb >= 20:
                self.calculate_key_id(key)
                self.calculate_sum(sum_qb)
                break
        return self.current_query

    def calculate_coef(self):
        '''Расчитывает коэффициент на основании данных о расходе коэффициента,
        в случае если потребления материалов не случилось для данных ремонтов,
        расчитывает на основании существующих коэффициентов'''
        sum_qua_bal = self.current_query["Quantity Balance"].sum(axis=0, skipna=False)
        sum_c_con = self.current_query["C Consum"].sum(axis=0, skipna=False)
        if sum_c_con == 0 or sum_c_con is None:
            self.coef = self.current_query["C quantity"].mean(axis=0)
            if self.coef == 0:
                self.coef = 0.01
        else:
            self.coef = sum_c_con / sum_qua_bal
        return self.coef

    def calculate_key_id(self, key):
        self.key_id = key
        return self.key_id

    def calculate_date(self, date):
        self.date_id = date
        return self.date_id

    def calculate_sum(self, sum_qb):
        self.sum_id = sum_qb
        return self.sum_id

if __name__ == '__main__':
    pass
