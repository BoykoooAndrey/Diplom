# 1.1 L_k
def find_k1_for_lk(category, table):
    return float(table[category][3].value)


def find_k2_for_lk(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(14):
            if inp_arr[i].modif_vihicle == (table[j+1][0].value.lower()):
                res_arr[i] = float(table[j+1][3].value)
    return res_arr


def find_k3_for_lk(weather, table):
    for i in range(10):
        if weather in (table[i+1][0].value.split(';')):
            return float(table[i+1][3].value)

# 1.2 L_to


def find_normat_l_to1(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(1, 10):
            if inp_arr[i].type_vihicle.lower() == (table[j][0].value.lower()):
                res_arr[i] = table[j][2].value
    return res_arr


def find_normat_l_to2(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(1, 10):
            if inp_arr[i].type_vihicle.lower() == (table[j][0].value.lower()):
                res_arr[i] = table[j][3].value
    return res_arr


# 1.5 n_g
def find_begin_end_index(car):
    if car.type_vihicle == 'автомобили легковые':
        begin_index = 2
        end_index = 5
    elif car.type_vihicle == 'автобусы':
        begin_index = 5
        end_index = 10
    elif car.type_vihicle == 'автомобили грузовые':
        begin_index = 10
        end_index = 18
    elif car.type_vihicle == 'полуприцепы':
        begin_index = 23
        end_index = 25
    elif car.type_vihicle == 'полуприцепы тяжеловозы':
        begin_index = 26
        end_index = 27
    return [begin_index, end_index]


def find_inTOTR_inKR(inp_arr, table):
    res_arr = []
    for i in range(len(inp_arr)):
        inTOTR_inKR = []
        indexes = find_begin_end_index(inp_arr[i])
        for j in range(indexes[0], indexes[1]):
            temp_list = (table[j][2].value).split(';')
            if inp_arr[i].loadCapacity > float(temp_list[0]) and inp_arr[i].loadCapacity <= float(temp_list[1]):
                inTOTR_inKR.append(float(table[j][4].value))
                inTOTR_inKR.append(float(table[j][5].value))
                break
        res_arr.append(inTOTR_inKR)
    return res_arr


def find_k2_for_TO_TR(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(14):
            if inp_arr[i].modif_vihicle == (table[j+1][0].value.lower()):
                res_arr[i] = float(table[j+1][2].value)
    return res_arr

# 1.10 Metod org


def metod_EO(N_EO_smen_day_sum):
    if N_EO_smen_day_sum >= 100:
        return 'Поточный метод организации обслуживания'
    else:
        return 'Метод организации на универсальных постах'


def metod_to1(N_1_day_sum):
    if N_1_day_sum >= 12:
        return 'Поточный метод организации обслуживания'
    else:
        return 'Метод организации на универсальных постах'


def metod_to2(N_2_day_sum):
    if N_2_day_sum >= 5:
        return 'Поточный метод организации обслуживания'
    else:
        return 'Метод организации на универсальных постах'


def metod_D1(metod_to1, metod_to2):
    if (metod_to1 == 'Поточный метод организации обслуживания') and (metod_to2 == 'Поточный метод организации обслуживания'):
        return 'Поточный метод организации обслуживания'
    elif (metod_to1 == 'Поточный метод организации обслуживания') and (metod_to2 == 'Метод организации на универсальных постах'):
        return 'На линии ТО-1'
    elif (metod_to1 == 'Метод организации на универсальных постах') and (metod_to2 == 'Метод организации на универсальных постах'):
        return 'Метод организации на универсальных постах'

# 2.1 trudoemkost


def find_trud_normat(inp_arr, table):
    res_arr = []
    for i in range(len(inp_arr)):
        indexes = find_begin_end_index(inp_arr[i])
        arr_trud_normat = []
        for j in range(indexes[0], indexes[1]):
            temp_list = (table[j][2].value).split(';')
            if inp_arr[i].loadCapacity > float(temp_list[0]) and inp_arr[i].loadCapacity <= float(temp_list[1]):
                for k in range(6, 10):
                    arr_trud_normat.append(float(table[j][k].value))
        res_arr.append(arr_trud_normat)
    return res_arr


def find_k1_for_trud_TR(category, table):
    return float(table[category][2].value)


def find_k2_for_trud(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(14):
            if inp_arr[i].modif_vihicle == (table[j+1][0].value.lower()):
                res_arr[i] = float(table[j+1][1].value)
    return res_arr
# def find_k4_trud(quantity_vehicle, table):
#     for i in range(2, 20):
#         temp_list = (table[i][0].value).split(';')
#         if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle < float(temp_list[1]):
#             return float(table[i][1].value)


def find_k4_trud(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        for j in range(2, 20):
            temp_list = (table[j][0].value).split(';')
            if inp_arr[i].quantity >= float(temp_list[0]) and inp_arr[i].quantity < float(temp_list[1]):
                res_arr[i] = float(table[j][1].value)
    return res_arr
# 3.1 raspred rabot


def find_index_for_tab_17(car):
    if car.type_vihicle == 'автомобили легковые':
        index = 1
    elif car.type_vihicle == 'автобусы':
        index = 2
    elif car.type_vihicle == 'автомобили грузовые':
        index = 3
    elif car.type_vihicle == 'автомобили самосвалы внедорожные':
        index = 4
    elif car.type_vihicle == 'полуприцепы тяжеловозы':
        index = 5
    return index


def find_fractions(inp_arr, table, T_year, a, b):
    res_arr = []
    for i in range(len(inp_arr)):
        fractions_EO = []
        tmp_index = find_index_for_tab_17(inp_arr[i])
        for j in range(a, b):
            vid_rabot = str(table[j][0].value)
            procent_rabot_ot_all = float(table[j][tmp_index].value)
            trudoemkost_rabot = round(
                float(table[j][tmp_index].value) * T_year[i] / 100, 2)
            cell = [vid_rabot, procent_rabot_ot_all, trudoemkost_rabot]
            fractions_EO.append(cell)
        res_arr.append(fractions_EO)
    return res_arr


def find_fractions_TO_1(type_vihicle, table):
    fractions_TO_1 = []
    if type_vihicle == 'автомобили легковые':
        for i in range(14, 16):
            fractions_TO_1.append(float(table[i][1].value))
    if type_vihicle == 'автобусы':
        for i in range(14, 16):
            fractions_TO_1.append(float(table[i][2].value))
    if type_vihicle == 'автомобили грузовые':
        for i in range(14, 16):
            fractions_TO_1.append(float(table[i][3].value))
    return fractions_TO_1


def find_fractions_TO_2(type_vihicle, table):
    fractions_TO_2 = []
    if type_vihicle == 'автомобили легковые':
        for i in range(18, 20):
            fractions_TO_2.append(float(table[i][1].value))
    if type_vihicle == 'автобусы':
        for i in range(18, 20):
            fractions_TO_2.append(float(table[i][2].value))
    if type_vihicle == 'автомобили грузовые':
        for i in range(18, 20):
            fractions_TO_2.append(float(table[i][3].value))
    return fractions_TO_2

# 3 X TO X D


def days_job_for_to1(days_job):
    modes_to_1 = []
    if days_job == 255:
        modes_to_1.append(float(255))
        modes_to_1.append(float(1))
        modes_to_1.append(float(8))
    elif days_job == 305:
        modes_to_1.append(float(305))
        modes_to_1.append(float(1))
        modes_to_1.append(float(8))
    elif days_job == 365:
        modes_to_1.append(float(365))
        modes_to_1.append(float(1))
        modes_to_1.append(float(8))
    return modes_to_1


def days_job_for_to2(days_job):
    modes_to_2 = []
    if days_job == 255:
        modes_to_2.append(float(255))
        modes_to_2.append(float(1))
        modes_to_2.append(float(8))
    elif days_job == 305:
        modes_to_2.append(float(305))
        modes_to_2.append(float(1))
        modes_to_2.append(float(8))
    elif days_job == 365:
        modes_to_2.append(float(365))
        modes_to_2.append(float(1))
        modes_to_2.append(float(8))
    return modes_to_2


def days_job_for_D(days_job):
    modes_D = []
    if days_job == 255:
        modes_D.append(float(255))
        modes_D.append(float(1))
        modes_D.append(float(8))
    elif days_job == 305:
        modes_D.append(float(305))
        modes_D.append(float(1))
        modes_D.append(float(8))
    elif days_job == 365:
        modes_D.append(float(365))
        modes_D.append(float(1))
        modes_D.append(float(8))
    return modes_D


def days_job_for_EO(days_job, table):
    modes_EO = []

    if days_job <= 305:
        modes_EO.append(float(305))
        modes_EO.append(float(1))
        modes_EO.append(float(8))
    elif days_job == 357:
        modes_EO.append(float(357))
        modes_EO.append(float(2))
        modes_EO.append(float(7))
    elif days_job == 365:
        modes_EO.append(float(365))
        modes_EO.append(float(2))
        modes_EO.append(float(7))
    return modes_EO


def days_job_for_TR(days_job, table):
    modes_TR = []
    if days_job == 255:
        modes_TR.append(float(255))
        modes_TR.append(float(1))
        modes_TR.append(float(8))
    elif days_job == 305:
        modes_TR.append(float(305))
        modes_TR.append(float(1))
        modes_TR.append(float(8))
    elif days_job == 357:
        modes_TR.append(float(357))
        modes_TR.append(float(1))
        modes_TR.append(float(8))
    elif days_job == 365:
        modes_TR.append(float(365))
        modes_TR.append(float(2))
        modes_TR.append(float(7))
    return modes_TR

# skads

def find_k3_sklad(car, table):
    indexes = find_begin_end_index(car)
    for j in range(indexes[0], indexes[1]):
        temp_list = (table[j][2].value).split(';')
        if car.loadCapacity > float(temp_list[0]) and car.loadCapacity <= float(temp_list[1]):
            return (float(table[j][10].value))


#etalon
# 5 ТЕХНИКО-ЭКОНОМИЧЕСКОЕ ОБОСНОВАНИЕ ПРОЕКТНЫХ РЕШЕНИЙ
def     find_standart(type_vihicle, table):
    arr = []
    if type_vihicle == 'автомобили легковые':
        for i in range(2, 8):
            arr.append(table[i][1].value)
    if type_vihicle == 'автобусы':
        for i in range(2, 8):
            arr.append(table[i][2].value)
    if type_vihicle == 'автомобили грузовые':
        for i in range(2, 8):
            arr.append(table[i][3].value)
    if type_vihicle == 'автомобили-самосвалы внедорожные':
        for i in range(2, 8):
            arr.append(table[i][4].value)
    return arr

def find_k_1_standart(quantity_vehicle, table):
    arr = []
    for i in range(2, 20):
        temp_list = (table[i][0].value).split(';')
        if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle < float(temp_list[1]):
            for j in range(4, 9):
                arr.append(float(table[i-1][j].value))

    return arr


def find_k2_standart(car, table):
    indexes = find_begin_end_index(car)
    arr = []
    for i in range(indexes[0], indexes[1]):
        temp_list = (table[i][2].value).split(';')
        if car.loadCapacity > float(temp_list[0]) and car.loadCapacity <= float(temp_list[1]):
            for j in range(12, 18):
                arr.append(float(table[i][j].value))
            break
    return arr

def find_k3_standart(procent_pricep, table):
    procent_pricep = procent_pricep * 100
    arr = []
    for i in range(2, 7):
        temp_list = (table[i][0].value).split(';')
        if procent_pricep > float(temp_list[0]) and procent_pricep <= float(temp_list[1]):
            for j in range(1, 7):
                arr.append(float(table[i][j].value))
            break
    return arr



























# нахождение габаритный размеров
def find_length_vihecle(model_vihicle, table):
    for i in range(2, 15):
        if table[i][0].value == model_vihicle:
            return round((float(table[i][1].value))/1000, 2)


def find_width_vihecle(model_vihicle, table):
    for i in range(2, 15):
        if table[i][0].value == model_vihicle:
            return round((float(table[i][2].value))/1000, 2)


def find_f_ud_x(model_vihicle, table):
    for i in range(2, 15):
        if table[i][0].value == model_vihicle:
            return (table[i][5].value)

# нахождение коэфицентов К


def find_k1_for_l_to(category, table):
    return float(table[category][1].value)


def find_k1_sklad(day_km):
    if day_km > 0 and day_km <= 100:
        return 0.8
    elif day_km > 100 and day_km <= 150:
        return 0.85
    elif day_km > 150 and day_km <= 200:
        return 0.8
    elif day_km > 200 and day_km <= 250:
        return 1
    elif day_km > 250 and day_km <= 300:
        return 1.15
    elif day_km > 300 and day_km <= 350:
        return 1.25
    else:
        return 1.15





def find_k2_sklad(quantity_vehicle, table):
    for i in range(2, 20):
        temp_list = (table[i][0].value).split(';')
        if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle < float(temp_list[1]):
            return float(table[i][2].value)


def find_k3_for_l_to(weather, table):
    for i in range(10):
        if weather in (table[i+1][0].value.split(';')):
            return float(table[i+1][1].value)


def find_k3_for_trud_TR(weather, table):
    for i in range(10):
        if weather in (table[i+1][0].value.split(';')):
            return float(table[i+1][2].value)


def find_k4_standart(day_km, table):
    arr = []
    for i in range(2, 8):
        temp_list = (table[i][0].value).split(';')
        if day_km > float(temp_list[0]) and day_km <= float(temp_list[1]):
            for j in range(1, 6):
                arr.append(table[i][j].value)
    return arr


def find_k5_trud(inp_arr, table):
    res_arr = [i for i in range(len(inp_arr))]
    for i in range(len(inp_arr)):
        if inp_arr[i].metod_keeping == 'open':
            res_arr[i] = 1
        else:
            res_arr[i] = 0.9
    return res_arr
    
    


def find_k5_sklad(category, table):
    return float(table[category][4].value)


def find_k6_standart(category, table):
    arr = []
    for i in range(6, 11):
        arr.append(float(table[category][i].value))
    return arr


def find_k7_standart(weather, table):
    arr = []
    for i in range(2, 8):
        if weather in (table[i][0].value.split(';')):
            for j in range(6, 11):
                arr.append(float(table[i][j].value))
    return arr
# 3.2.1. Корректирование норм пробегов до ТО и КР


def find_class_lkn(type_vihicle, characteristic, table):
    class_and_lkn = []
    if type_vihicle == 'автомобили легковые':
        for i in range(2, 5):
            temp_list = (table[i][2].value).split(';')
            if characteristic > float(temp_list[0]) and characteristic <= float(temp_list[1]):
                class_and_lkn.append(table[i][1].value)
                class_and_lkn.append(float(table[i][3].value))
    if type_vihicle == 'автобусы':
        for i in range(5, 10):
            temp_list = (table[i][2].value).split(';')
            if characteristic > float(temp_list[0]) and characteristic <= float(temp_list[1]):
                class_and_lkn.append(table[i][1].value)
                class_and_lkn.append(float(table[i][3].value))
    if type_vihicle == 'автомобили грузовые':
        for i in range(10, 18):
            temp_list = (table[i][2].value).split(';')
            if characteristic > float(temp_list[0]) and characteristic <= float(temp_list[1]):
                class_and_lkn.append(table[i][1].value)
                class_and_lkn.append(float(table[i][3].value))
    return class_and_lkn

# 3.2.3. Расчет количества ТО и КР (списаний) на весь парк за год


# 3.2.5. Расчет суточной производственной программы по видам ЕО и ТО


# 3.2.8. Корректирование нормативных трудоемкостей


# 3.2.9. Расчет годовых объемов работ по ЕО, ТО, Д, ТР


def find_fraction_d1_in_to1(type_vihicle, table):
    if type_vihicle == 'автомобили легковые':
        return (float(table[14][1].value))
    if type_vihicle == 'автобусы':
        return (float(table[14][2].value))
    if type_vihicle == 'автомобили грузовые':
        return (float(table[14][3].value))


def find_fraction_d2_in_to2(type_vihicle, table):
    if type_vihicle == 'автомобили легковые':
        return (float(table[18][1].value))
    if type_vihicle == 'автобусы':
        return (float(table[18][2].value))
    if type_vihicle == 'автомобили грузовые':
        return (float(table[18][3].value))


def find_fraction_d1_in_tr(type_vihicle, table):
    if type_vihicle == 'автомобили легковые':
        return (float(table[23][1].value))
    if type_vihicle == 'автобусы':
        return (float(table[23][2].value))
    if type_vihicle == 'автомобили грузовые':
        return (float(table[23][3].value))


def find_fraction_d2_in_tr(type_vihicle, table):
    if type_vihicle == 'автомобили легковые':
        return (float(table[24][1].value))
    if type_vihicle == 'автобусы':
        return (float(table[24][2].value))
    if type_vihicle == 'автомобили грузовые':
        return (float(table[24][3].value))

# 3.2.11. Распределение годовых объемов работ по производственным зонам и участкам (цехам)


def find_fractions_post(type_vihicle, table):
    fractions_post = []
    if type_vihicle == 'автомобили легковые':
        for i in range(23, 28):
            fractions_post.append(float(table[i][1].value))
    if type_vihicle == 'автобусы':
        for i in range(23, 28):
            fractions_post.append(float(table[i][2].value))
    if type_vihicle == 'автомобили грузовые':
        for i in range(23, 28):
            fractions_post.append(float(table[i][3].value))
    return fractions_post


def find_fractions_uch(type_vihicle, table):
    fractions_uch = []
    if type_vihicle == 'автомобили легковые':
        for i in range(30, 44):
            fractions_uch.append(float(table[i][1].value))
    if type_vihicle == 'автобусы':
        for i in range(30, 44):
            fractions_uch.append(float(table[i][2].value))
    if type_vihicle == 'автомобили грузовые':
        for i in range(30, 44):
            fractions_uch.append(float(table[i][3].value))
    return fractions_uch


def find_fractions_T_pp(type_vihicle, table):
    fractions_T_pp = []
    for i in range(52, 59):
        fractions_T_pp.append(float(table[i][1].value))
    return fractions_T_pp


def find_fractions_T_sam(type_vihicle, table):
    fractions_T_sam = []
    for i in range(61, 69):
        fractions_T_sam.append(float(table[i][1].value))
    return fractions_T_sam


def find_koef_plotn_uch(name_uch, table):
    for i in range(1, 4):
        for tmp in ((table[i][0].value).split(",")):
            if tmp == name_uch.lower():
                return table[i][1].value

# нахождение рабочих на ТО-1 и на ТО-2


def find_distance_between_cars(type_vihicle):
    if type_vihicle == 'автомобили грузовые':
        return 1.5
    elif type_vihicle == 'автомобили легковые':
        return 1.2
    elif type_vihicle == 'автобусы':
        return 2

# коэффициента использования рабочего времени поста


def find_n_2_for_to_2(c, table):
    return float(table[7][int(c)].value)


def find_n_2_for_EO(c, table):
    arr = [float(table[3][int(c)].value), float(table[4][int(c)].value)]
    return arr


def find_n_2_for_TR(c, table):
    return float(table[11][int(c)].value)

# Расчет ЕО


# Значения коэффициентов неравномерности поступления автомобилей на рабочие посты (φ)
def find_fi(kakya_zona, quantity_vehicle, c, table):
    if kakya_zona == 'EO' or kakya_zona == 'TR':
        temp1 = int(3)  # номер строки
    else:
        temp1 = int(4)
    if quantity_vehicle < 100:
        columns = [1, 2]  # номер столбца
    elif quantity_vehicle >= 100 and quantity_vehicle <= 300:
        columns = [3, 4]
    elif quantity_vehicle > 300 and quantity_vehicle <= 500:
        columns = [5, 6]
    elif quantity_vehicle > 500 and quantity_vehicle <= 1000:
        columns = [7, 8]
    elif quantity_vehicle > 1000 and quantity_vehicle <= 2000:
        columns = [9, 10]
    elif quantity_vehicle > 2000:
        columns = [11, 12]
    if c == 1:
        return float(table[temp1][columns[0]].value)
    else:
        return float(table[temp1][columns[1]].value)


# продолжительность «пикового» возврата подвижного состава в течение суток на АТП
def find_T_vozvr(type_vihicle, quantity_vehicle, table):

    if type_vihicle == 'автомобили легковые':
        for i in range(2, 20):
            temp_list = (table[i][0].value).split(';')
            if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle <= float(temp_list[1]):
                return float(table[i][1].value)
    if type_vihicle == 'автобусы':
        for i in range(2, 20):
            temp_list = (table[i][0].value).split(';')
            if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle <= float(temp_list[1]):
                return float(table[i][2].value)
    if type_vihicle == 'автомобили грузовые':
        for i in range(2, 20):
            temp_list = (table[i][0].value).split(';')
            if quantity_vehicle >= float(temp_list[0]) and quantity_vehicle <= float(temp_list[1]):
                return float(table[i][3].value)


def find_fractions_post_all(type_vihicle, table):
    if type_vihicle == 'автомобили легковые':
        return float(table[27][1].value)
    if type_vihicle == 'автобусы':
        return float(table[27][2].value)
    if type_vihicle == 'автомобили грузовые':
        return float(table[27][3].value)


# 3.2.21. Расчет площадей производственных участков
def find_f_1_f_2_for_regions(table):
    Arr = []
    for i in range(19, 33):
        if table[i][4].value != 0:
            arr = []
            arr.append(table[i][0].value)
            arr.append(table[i][1].value)
            arr.append(table[i][2].value)
            arr.append(table[i][3].value)
            arr.append(table[i][4].value)

            Arr.append(arr)
    return Arr

# 3.2.23. Расчет площадей складских помещений


def find_f_sklad_k_4(type_vihicle, table):
    Arr = []
    if type_vihicle == 'автомобили легковые':
        for i in range(2, 13):
            arr = []
            arr.append(table[i][0].value)
            arr.append(table[i][1].value)
            arr.append(table[i][6].value)
            Arr.append(arr)
    if type_vihicle == 'автобусы':
        for i in range(2, 13):
            arr = []
            arr.append(table[i][0].value)
            arr.append(table[i][2].value)
            arr.append(table[i][6].value)
            Arr.append(arr)
    if type_vihicle == 'автомобили грузовые':
        for i in range(2, 13):
            arr = []
            arr.append(table[i][0].value)
            arr.append(table[i][3].value)
            arr.append(table[i][6].value)
            Arr.append(arr)
    if type_vihicle == 'полуприцепы тяжеловозы':
        for i in range(2, 13):
            arr = []
            arr.append(table[i][0].value)
            arr.append(table[i][4].value)
            arr.append(table[i][6].value)
            Arr.append(arr)
    return Arr



