import lib

data_file_name = "RU05_query.xlsx"
user_file_name = input("Please enter a file name: ")

data_list = lib.Data_list(data_file_name)
mat_list = data_list.create_materials()
zparts_list = data_list.create_zparts()

user_data = lib.User_data(user_file_name)
user_list = user_data.create_df()
prdinfo = user_data.create_prdinfo()

product = lib.Product_obj(0, prdinfo, mat_list)
product.prepare_objdata()
product.create_obj()
product_obj = product.rename_obj()

for i in range(len(user_list)):
    component = lib.Component_obj(i, user_list, mat_list)
    component.prepare_objdata()
    component.create_obj()
    component_obj = component.rename_obj()
    complete_df = lib.Complete_df(zparts_list, component_obj, product_obj)
    complete_df.join_obj()
    complete_df.create_query()
    complete_df.calculate_coef()
    user_data.create_coef_array(complete_df.coef, complete_df.key_id,
                                complete_df.date_id, complete_df.sum_id)
user_data.coef_arr_concat()
user_data.print_df()

