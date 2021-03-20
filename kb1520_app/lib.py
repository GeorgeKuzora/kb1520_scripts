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
            self.materials_df = pd.read_excel(xlsx, "materials_combined", na_values=["NA"], index_col=0, usecols="A:E")
        return self.materials_df

    def create_zparts(self):
        '''Создает датафрейм из листа Эксель в котором содержаться данные о
        произведенных ремонтах и расходе материалов в этих ремонтах.'''
        with pd.ExcelFile(self.basic_file) as xlsx:
            self.zparts_df = pd.read_excel(xlsx, "zparts_combined", na_values=["NA"], usecols="A:U")
        return self.zparts_df


class User_data:
    '''Класс для загрузки и обработки данных загружаеммых пользователем,
    а именно ВОМ материала и данные о продукте занесенные в пользовательский файл.
    Требует очистки данных пользователя'''

    def __init__(self, file_name):
        '''Атрибут имя файла .xlsx где храняться данные разделенные по вкладкам,
        сформированныи при промощи Powerquery'''
        self.file_name = file_name
        self.coef_array = []

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

    def create_coef_array(self, coef):
        '''Добавляет расчитанный для объекта компонента, входящего в лист с
        пользовательскими данными, через который идет итерация, в лист для
        последующего добвавления к данным'''
        self.coef_array.append(coef)

    def coef_arr_concat(self):
        '''Присоединяет лист к добавленными коэффициентами к основному дата фрейму
        пользовательских данных'''
        coef_arr_s = pd.Series(self.coef_array, name="Coef")
        self.user_df = pd.concat([self.user_df, coef_arr_s], axis=1)
        return self.user_df

    def print_df(self):
        '''Записывает пользовательский датафрейм в новый эксель файл'''
        self.user_df.to_excel("true_" + self.file_name, sheet_name="BOM")


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
                            "Material type": "Component material type",
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
                            "Material type": "Product material type",
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
                 4: ("Component", "Product material type"),
                 5: ("Component", "Product COC"),
                 6: ("Component description", "Service material"),
                 7: ("Component description", "Product"),
                 8: ("Component description", "Product description"),
                 9: ("Component description", "Product material type"),
                 10: ("Component description", "Product COC"),
                 11: ("Component material type", "Product COC"),
                 12: ("Component COC", "Product COC")
                }

    def __init__(self, zparts_df, component_obj, product_obj):
        '''Атрибуты:
            Данные о ремонтах: данные из датафрейма zparts_df,
            Объект компонента: объект компонента для каждого из материалов в
            пользовательских данных,
            Объект продукта: объект продукта для которого производиться ремонт'''
        self.zparts_df = zparts_df
        self.component_obj = component_obj
        self.product_obj = product_obj

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
            #current_frame = self.joined_obj[:, value]
            component_id = self.joined_obj.at[0, value[0]]
            product_id = self.joined_obj.at[0, value[1]]
            #if component_id is None:
            #   self.current_query = self.zparts_df[(self.zparts_df[value[1]] == product_id)]
            self.current_query = self.zparts_df[(self.zparts_df[value[0]] == component_id) & (self.zparts_df[value[1]] == product_id)]
            sum_qb = self.current_query["Quantity Balance"].sum()
            if sum_qb >= 20:
                break
        return self.current_query

    def calculate_coef(self):
        '''Расчитывает коэффициент на основании данных о расходе коэффициента,
        в случае если потребления материалов не случилось для данных ремонтов,
        расчитывает на основании существующих коэффициентов'''
        if self.current_query["C Consum"].sum(axis=0, skipna=False) == 0:
            self.coef = self.current_query["C quantity"].min(axis=0, skipna=False)
        else:
            self.coef = self.current_query["Prob_C"].mean(axis=0, skipna=False)
        return self.coef




if __name__ == '__main__':
    pass
