import statistics
import car
import random
import openpyxl
import math

# vid_to_tr == 1 - TO-1; 2 - TO-2;  3 - TR

uaz_3163 = car.Vehicle("УАЗ-3163", 29, 2.7, 20, 'автомобили легковые', 'базовый автомобиль', 122.24, 300000, 4.647, 1.929)
kamaz_43502 = car.Vehicle("КАМАЗ-43502", 124, 4, 105, 'автомобили грузовые', 'базовый автомобиль', 89.85, 300000, 7.570, 2.5)
kamaz_43118 = car.Vehicle("КАМАЗ-43118", 144, 10, 93, 'автомобили грузовые', 'седельные тягачи', 107, 300000, 8.835, 2.5)
politrans = car.Vehicle("ПОЛИТРАНС-94163", 61, 40, 32, 'полуприцепы тяжеловозы', 'полуприцепы', 113.01, 250000, 15.130, 2.550)
cars = [uaz_3163, kamaz_43502, kamaz_43118, politrans]
LEN_CAR = len(cars)

L_g=[28054.08, 21995.28, 25647.9, 27953.02]
for i in range(LEN_CAR):
    L_g[i] = L_g[i] * cars[i].quantity








def find_k_3(weather):
    for i in range(18, 24):
        if weather in (table_4[i+1][0].value.split(';')):
            return float(table_4[i+1][1].value)


def find_length_proetk_zd(S_proetk_zd):
    for i in range(6, 14):
        if S_proetk_zd % i == 0:
            return i


init_data = openpyxl.open('init_data.xlsx', read_only=True)
table_1 = init_data.worksheets[0]
table_2 = init_data.worksheets[1]
table_3 = init_data.worksheets[2]
table_4 = init_data.worksheets[3]

arr_headings_output_tregub = []
arr_value_output_tregub = []

for i in range(LEN_CAR):
    arr_headings_output_tregub.append(f'cars{i}')
    arr_value_output_tregub.append(cars[i].name)

VARIANT = 10


# 1. РАСЧЕТ ПЛАНОВОЙ ЧИСЛЕННОСТИ РАБОТНИКОВ
# 1.2. Расчет численности младшего обслуживающего персонала
N_to1 = table_2[26][13].value
N_to2 = table_2[27][13].value
N_tr = table_2[28][13].value
N_rr = table_2[29][13].value
N_vr = table_2[30][13].value
N_mon = 0.02*(N_rr + N_vr)
if N_mon < 1:
    N_mon = math.ceil(0.02*(N_rr + N_vr))
else:
    N_mon = round(0.02*(N_rr + N_vr))

# 1.3. Определение численности руководителей, специалистов и служащих
N_rss = round(0.2 * (N_vr + N_rr + N_mon))

# 1.4. Общая численность работников
N_all = N_rr + N_vr + N_rss + N_mon


def append_1():
    arr_headings_output_tregub.append('N_to1')
    arr_value_output_tregub.append(N_to1)
    arr_headings_output_tregub.append('N_to2')
    arr_value_output_tregub.append(N_to2)
    arr_headings_output_tregub.append('N_tr')
    arr_value_output_tregub.append(N_tr)
    arr_headings_output_tregub.append('N_rr')
    arr_value_output_tregub.append(N_rr)
    arr_headings_output_tregub.append('N_vr')
    arr_value_output_tregub.append(N_vr)
    arr_headings_output_tregub.append('N_mon')
    arr_value_output_tregub.append(N_mon)
    arr_headings_output_tregub.append('N_rss')
    arr_value_output_tregub.append(N_rss)
    arr_headings_output_tregub.append('N_all')
    arr_value_output_tregub.append(N_all)
append_1()

# 2. РАСЧЕТ ПЛАНОВОГО ФОНДА ЗАРАБОТНОЙ ПЛАТЫ
# 2.1. Определение средней тарифной ставки по видам воздействий (ТО-1, ТО-2, ТР)
tarif_3_raz = table_1[VARIANT + 1][4].value
tarif_4_raz = table_1[VARIANT + 1][5].value
tarif_5_raz = table_1[VARIANT + 1][6].value
tarif_6_raz = table_1[VARIANT + 1][7].value

workers_to_1 = []  # В массиве идет по порядку 3,4,5,6 разряд
workers_to_2 = []  # В массиве идет по порядку 3,4,5,6 разряд
workers_TR = []  # В массиве идет по порядку 3,4,5,6 разряд
for i in range(9, 13):
    workers_to_1.append(int(table_2[26][i].value))
    workers_to_2.append(int(table_2[27][i].value))
    workers_TR.append(int(table_2[28][i].value))
sred_tarif_to_1 = round((workers_to_1[0] * tarif_3_raz + workers_to_1[1] *
                         tarif_4_raz + workers_to_1[2] * tarif_5_raz + workers_to_1[3] * tarif_6_raz) / sum(workers_to_1), 2)
sred_tarif_to_2 = round((workers_to_2[0] * tarif_3_raz + workers_to_2[1] *
                         tarif_4_raz + workers_to_2[2] * tarif_5_raz + workers_to_2[3] * tarif_6_raz) / sum(workers_to_2), 2)
sred_tarif_tr = round((workers_TR[0] * tarif_3_raz + workers_TR[1] *
                       tarif_4_raz + workers_TR[2] * tarif_5_raz + workers_TR[3] * tarif_6_raz) / sum(workers_TR), 2)

# 2.2 Определение заработной платы ремонтных рабочих по тарифу по видам воздействий
# 2.2.1 Фонд заработной платы ремонтных рабочих
T_to_1 = float(table_1[VARIANT + 1][2].value)
T_to_2 = float(table_1[VARIANT + 1][3].value)
T_tr = float(table_1[VARIANT + 1][1].value)
FZP_to_1 = round(sred_tarif_to_1*T_to_1, 2)
FZP_to_2 = round(sred_tarif_to_2*T_to_2, 2)
FZP_tr = round(sred_tarif_tr*T_tr, 2)

# 2.2.2. Общий фонд заработной платы ремонтных рабочих по тарифу
FZP_all = FZP_to_1 + FZP_to_2 + FZP_tr

# 2.2.3. Средняя часовая тарифная ставка ремонтных рабочих
T_all = T_to_1 + T_to_2 + T_tr
sred_tarif_rem_workers = round(FZP_all / T_all)

# 2.3. Общий фонд заработной платы руководителей, специалистов и служащих (РСС)

Oklad_rss = float(table_1[VARIANT + 1][34].value)
ZP_rss = Oklad_rss * 12 * N_rss

# 2.3.1. Премия РСС
prize_rss = round(0.4 * ZP_rss, 2)

# 2.3.2. Основная заработная плата РСС
OZP_rss = round(ZP_rss + prize_rss, 2)

# 2.3.3. Дополнительная заработная плата РСС
vacation_days = 28  # Количество дней отпуска
calendar_days = 365  # Количество календарных дней
weekend_days = 118  # Количество выходных дней
weather = table_1[VARIANT + 1][32].value
if weather == 'холодные':
    vacation_days = 44
if weather == 'очень холодные':
    vacation_days = 52
DZP_rss = round(
    (vacation_days/(calendar_days - weekend_days - vacation_days))*OZP_rss, 2)

# 2.3.4. Единовременные поощрительные выплаты РСС
EPV_rss = round(0.02 * OZP_rss, 2)

# 2.3.5. Общий фонд заработной платы РСС
# умеренные # умеренно-теплые # умеренно-теплые влажные # теплые влажные # жаркий сухой # очень жаркий сухой
# умеренно-холодные RK 1.15 # холодные RK 1.7 SN SN 0.5 # очень холодные RK 2 SN 2
if weather == 'умеренно-холодные':
    k = 1.15
elif weather == 'холодные':
    k = 2.2
elif weather == 'очень холодные':
    k = 2.5
else:
    k = 1
OFZP_rss = round((OZP_rss + DZP_rss) * k + EPV_rss, 2)

# 2.3.6. Среднемесячная заработная плата РСС
ZP_mouth_rss = round(OFZP_rss/(12*N_rss), 2)


def append_2_1():

    arr_headings_output_tregub.append('tarif_3_raz')
    arr_value_output_tregub.append(tarif_3_raz)
    arr_headings_output_tregub.append('tarif_4_raz')
    arr_value_output_tregub.append(tarif_4_raz)
    arr_headings_output_tregub.append('tarif_5_raz')
    arr_value_output_tregub.append(tarif_5_raz)
    arr_headings_output_tregub.append('tarif_6_raz')
    arr_value_output_tregub.append(tarif_6_raz)
    arr_headings_output_tregub.append('workers_to_1_3_raz')
    arr_value_output_tregub.append(workers_to_1[0])
    arr_headings_output_tregub.append('workers_to_1_4_raz')
    arr_value_output_tregub.append(workers_to_1[1])
    arr_headings_output_tregub.append('workers_to_1_5_raz')
    arr_value_output_tregub.append(workers_to_1[2])
    arr_headings_output_tregub.append('workers_to_1_6_raz')
    arr_value_output_tregub.append(workers_to_1[3])
    arr_headings_output_tregub.append('sred_tarif_to_1')
    arr_value_output_tregub.append(sred_tarif_to_1)
    arr_headings_output_tregub.append('workers_to_2_3_raz')
    arr_value_output_tregub.append(workers_to_2[0])
    arr_headings_output_tregub.append('workers_to_2_4_raz')
    arr_value_output_tregub.append(workers_to_2[1])
    arr_headings_output_tregub.append('workers_to_2_5_raz')
    arr_value_output_tregub.append(workers_to_2[2])
    arr_headings_output_tregub.append('workers_to_2_6_raz')
    arr_value_output_tregub.append(workers_to_2[3])
    arr_headings_output_tregub.append('sred_tarif_to_2')
    arr_value_output_tregub.append(sred_tarif_to_2)
    arr_headings_output_tregub.append('workers_TR_3_raz')
    arr_value_output_tregub.append(workers_TR[0])
    arr_headings_output_tregub.append('workers_TR_4_raz')
    arr_value_output_tregub.append(workers_TR[1])
    arr_headings_output_tregub.append('workers_TR_5_raz')
    arr_value_output_tregub.append(workers_TR[2])
    arr_headings_output_tregub.append('workers_TR_6_raz')
    arr_value_output_tregub.append(workers_TR[3])
    arr_headings_output_tregub.append('sred_tarif_TR')
    arr_value_output_tregub.append(sred_tarif_tr)
append_2_1()

def append_2_2():
    arr_headings_output_tregub.append('T_to_1')
    arr_value_output_tregub.append(T_to_1)
    arr_headings_output_tregub.append('T_to_2')
    arr_value_output_tregub.append(T_to_2)
    arr_headings_output_tregub.append('T_tr')
    arr_value_output_tregub.append(T_tr)
    arr_headings_output_tregub.append('FZP_to_1')
    arr_value_output_tregub.append(FZP_to_1)
    arr_headings_output_tregub.append('FZP_to_2')
    arr_value_output_tregub.append(FZP_to_2)
    arr_headings_output_tregub.append('FZP_tr')
    arr_value_output_tregub.append(FZP_tr)
    arr_headings_output_tregub.append('FZP_all')
    arr_value_output_tregub.append(FZP_all)
    arr_headings_output_tregub.append('T_all')
    arr_value_output_tregub.append(T_all)
    arr_headings_output_tregub.append('sred_tarif_rem_workers')
    arr_value_output_tregub.append(sred_tarif_rem_workers)
append_2_2()

def append_2_3():
    arr_headings_output_tregub.append('Oklad_rss')
    arr_value_output_tregub.append(Oklad_rss)
    arr_headings_output_tregub.append('ZP_rss')
    arr_value_output_tregub.append(ZP_rss)
    arr_headings_output_tregub.append('prize_rss')
    arr_value_output_tregub.append(prize_rss)
    arr_headings_output_tregub.append('OZP_rss')
    arr_value_output_tregub.append(OZP_rss)
    arr_headings_output_tregub.append('vacation_days')
    arr_value_output_tregub.append(vacation_days)
    arr_headings_output_tregub.append('calendar_days')
    arr_value_output_tregub.append(calendar_days)
    arr_headings_output_tregub.append('weekend_days')
    arr_value_output_tregub.append(weekend_days)
    arr_headings_output_tregub.append('DZP_rss')
    arr_value_output_tregub.append(DZP_rss)
    arr_headings_output_tregub.append('EPV_rss')
    arr_value_output_tregub.append(EPV_rss)
    arr_headings_output_tregub.append('k')
    arr_value_output_tregub.append(k)
    arr_headings_output_tregub.append('FZP_rss')
    arr_value_output_tregub.append(OFZP_rss)
    arr_headings_output_tregub.append('ZP_mouth_rss')
    arr_value_output_tregub.append(ZP_mouth_rss)
append_2_3()

# 2.4 Общий фонд заработной платы вспомогательных рабочих и младшего
OFZP_vrimon = round(0.3*FZP_all, 2)
# 2.4.1. Среднемесячная заработная плата вспомогательных рабочих и МОП
ZP_vrimon_mes = round(OFZP_vrimon / (12 * (N_vr + N_mon)), 2)
# 2.5. Общий фонд заработной платы РСС и МОП
OFZP_rss_vr_mon = OFZP_rss + OFZP_vrimon
# 2.5.1.Отчисление в соц. фонды
SOC_OTC_rss_vr_mon = round(OFZP_rss_vr_mon * 0.3, 2)
# 2.6. Доплата за условия труда по видам воздействий
N_vred_to1 = N_to1 * 0.05
if N_vred_to1 < 1:
    N_vred_to1 = math.ceil(N_to1 * 0.05)
else:
    N_vred_to1 = round(N_to1 * 0.05)
N_vred_to2 = N_to2 * 0.1
if N_vred_to2 < 1:
    N_vred_to2 = math.ceil(N_to2 * 0.1)
else:
    N_vred_to2 = round(N_to2 * 0.1)
N_vred_tr = N_tr * 0.1
if N_vred_tr < 1:
    N_vred_tr = math.ceil(N_tr * 0.1)
else:
    N_vred_tr = round(N_tr * 0.1)

sred_tarif_to_1_mes = sred_tarif_to_1 * 173.1
sred_tarif_to_2_mes = sred_tarif_to_2 * 173.1
sred_tarif_tr_mes = sred_tarif_tr * 173.1

D_vred_to1 = round((sred_tarif_to_1_mes * 12 * N_vred_to1 * 12)/100, 2)
D_vred_to2 = round((sred_tarif_to_2_mes * 12 * N_vred_to2 * 12)/100, 2)
D_vred_tr = round((sred_tarif_tr_mes * 12 * N_vred_tr * 12)/100, 2)
D_vred_all = D_vred_to1 + D_vred_to2 + D_vred_tr

# Общая сумма доплат за работу в ночные смены
N_vc = round(N_rr*0.07)
D_vc = round((20 * sred_tarif_rem_workers * 2 * 50 * N_vc)/100, 2)
D_ut_all = D_vred_to1 + D_vred_to2 + D_vred_tr + D_vc

# 2.7. Общий фонд заработной платы по тарифу
OFZP_tar = FZP_all + D_ut_all + OFZP_rss + OFZP_vrimon

# 2.8. Расчет сдельного расценка за 1 автомобиле-день работы автомобиля
AH_r = table_1[VARIANT + 1][8].value * table_1[VARIANT + 1][9].value * int(((table_1[VARIANT + 1][33].value).split(";"))[0])
C_cd = round(OFZP_tar/AH_r, 2)
# table 2.1
K_ud_to1 = round(FZP_to_1/FZP_all, 2)
K_ud_to2 = round(FZP_to_2/FZP_all, 2)
K_ud_tr = round(FZP_tr/FZP_all, 2)
OFZP_s_uchetom_D_to1 = FZP_to_1 + D_vred_to1
OFZP_s_uchetom_D_to2 = FZP_to_2 + D_vred_to2
OFZP_s_uchetom_D_tr = FZP_tr + D_vred_tr
OFZP_s_uchetom_D_all = OFZP_s_uchetom_D_to1 + \
    OFZP_s_uchetom_D_to2 + OFZP_s_uchetom_D_tr
C_cd_to1 = round(C_cd * K_ud_to1, 2)
C_cd_to2 = round(C_cd * K_ud_to2, 2)
C_cd_tr = round(C_cd * K_ud_tr, 2)


def append_2_2():
    arr_headings_output_tregub.append('OFZP_vrimon')
    arr_value_output_tregub.append(OFZP_vrimon)
    arr_headings_output_tregub.append('ZP_vrimon_mes')
    arr_value_output_tregub.append(ZP_vrimon_mes)
    arr_headings_output_tregub.append('OFZP_rss_vr_mon')
    arr_value_output_tregub.append(OFZP_rss_vr_mon)
    arr_headings_output_tregub.append('SOC_OTC_rss_vr_mon')
    arr_value_output_tregub.append(SOC_OTC_rss_vr_mon)
    arr_headings_output_tregub.append('N_vred_to1')
    arr_value_output_tregub.append(N_vred_to1)
    arr_headings_output_tregub.append('N_vred_to2')
    arr_value_output_tregub.append(N_vred_to2)
    arr_headings_output_tregub.append('N_vred_tr')
    arr_value_output_tregub.append(N_vred_tr)
    arr_headings_output_tregub.append('sred_tarif_to_1_mes')
    arr_value_output_tregub.append(sred_tarif_to_1_mes)
    arr_headings_output_tregub.append('sred_tarif_to_2_mes')
    arr_value_output_tregub.append(sred_tarif_to_2_mes)
    arr_headings_output_tregub.append('sred_tarif_tr_mes')
    arr_value_output_tregub.append(sred_tarif_tr_mes)
    arr_headings_output_tregub.append('D_vred_to1')
    arr_value_output_tregub.append(D_vred_to1)
    arr_headings_output_tregub.append('D_vred_to2')
    arr_value_output_tregub.append(D_vred_to2)
    arr_headings_output_tregub.append('D_vred_tr')
    arr_value_output_tregub.append(D_vred_tr)
    arr_headings_output_tregub.append('N_vc')
    arr_value_output_tregub.append(N_vc)
    arr_headings_output_tregub.append('D_vc')
    arr_value_output_tregub.append(D_vc)
    arr_headings_output_tregub.append('D_ut_all')
    arr_value_output_tregub.append(D_ut_all)
    arr_headings_output_tregub.append('OFZP_tar')
    arr_value_output_tregub.append(OFZP_tar)
    arr_headings_output_tregub.append('AH_r')
    arr_value_output_tregub.append(AH_r)
    arr_headings_output_tregub.append('C_cd')
    arr_value_output_tregub.append(C_cd)
    arr_headings_output_tregub.append('K_ud_to1')
    arr_value_output_tregub.append(K_ud_to1)
    arr_headings_output_tregub.append('K_ud_to2')
    arr_value_output_tregub.append(K_ud_to2)
    arr_headings_output_tregub.append('K_ud_tr')
    arr_value_output_tregub.append(K_ud_tr)
    arr_headings_output_tregub.append('C_cd_to1')
    arr_value_output_tregub.append(C_cd_to1)
    arr_headings_output_tregub.append('C_cd_to2')
    arr_value_output_tregub.append(C_cd_to2)
    arr_headings_output_tregub.append('C_cd_tr')
    arr_value_output_tregub.append(C_cd_tr)
append_2_2()
# 2.9. Сдельная заработная плата за год по видам воздействий
ZP_cd_to1 = round(C_cd_to1 * AH_r, 2)
ZP_cd_to2 = round(C_cd_to2 * AH_r, 2)
ZP_cd_tr = round(C_cd_tr * AH_r, 2)
ZP_cd_all = round(ZP_cd_to1 + ZP_cd_to2 + ZP_cd_tr, 2)
# 2.9.1. Сумма премий из фонда оплаты труда
P_RFZP_to1 = round(ZP_cd_to1 * 0.5, 2)
P_RFZP_to2 = round(ZP_cd_to2 * 0.5, 2)
P_RFZP_tr = round(ZP_cd_tr * 0.5, 2)
P_RFZP_all = P_RFZP_to1 + P_RFZP_to2 + P_RFZP_tr
# table 2.2
P_RFZP_to1_ZP_cd_to1 = ZP_cd_to1 + P_RFZP_to1
P_RFZP_to2_ZP_cd_to2 = ZP_cd_to2 + P_RFZP_to2
P_RFZP_tr_ZP_cd_tr = ZP_cd_tr + P_RFZP_tr
P_RFZP_all_ZP_cd_all = ZP_cd_all + P_RFZP_all

# 2.9.2. Расчет доплат
C_rr_sred = round(sred_tarif_rem_workers * ((FZP_all + D_vred_all)/FZP_all), 2)
# Расчет коэффициента доплат
K_d = round(D_vc/ZP_cd_all, 4)
# Сумма доплат по видам воздействия
D_vc_to1 = round(ZP_cd_to1 * K_d, 2)
D_vc_to2 = round(ZP_cd_to2 * K_d, 2)
D_vc_tr = round(ZP_cd_tr * K_d, 2)
# 2.9.3. Основная заработная плата работников по видам воздействия
OZP_to1 = round(ZP_cd_to1 + P_RFZP_to1 + D_vc_to1, 2)
OZP_to2 = round(ZP_cd_to2 + P_RFZP_to2 + D_vc_to2, 2)
OZP_tr = round(ZP_cd_tr + P_RFZP_tr + D_vc_tr, 2)
# 2.9.4. Дополнительная заработная плата по видам воздействия
DZP_to1 = round(
    (vacation_days/(calendar_days - weekend_days - vacation_days))*OZP_to1, 2)
DZP_to2 = round(
    (vacation_days/(calendar_days - weekend_days - vacation_days))*OZP_to2, 2)
DZP_tr = round(
    (vacation_days/(calendar_days - weekend_days - vacation_days))*OZP_tr, 2)
# 2.9.5. Общий фонд заработной платы по видам воздействия
OFZP_to1 = round((OZP_to1 + DZP_to1)*k, 2)
OFZP_to2 = round((OZP_to2 + DZP_to2)*k, 2)
OFZP_tr = round((OZP_tr + DZP_tr)*k, 2)
# 2.9.6. Отчисление в соц. фонды по видам воздействия
SOC_OTC_to1 = round((OFZP_to1 * 30) / 100, 2)
SOC_OTC_to2 = round((OFZP_to2 * 30) / 100, 2)
SOC_OTC_tr = round((OFZP_tr * 30) / 100, 2)
# 2.9.7. Фонд заработной платы с отчислениями по видам воздействия
OFZP_all_with_otc_to1 = OFZP_to1 + SOC_OTC_to1
OFZP_all_with_otc_to2 = OFZP_to2 + SOC_OTC_to2
OFZP_all_with_otc_tr = OFZP_tr + SOC_OTC_tr


def append_2_9():
    arr_headings_output_tregub.append('D_vred_all')
    arr_value_output_tregub.append(D_vred_all)
    arr_headings_output_tregub.append('ZP_cd_to1')
    arr_value_output_tregub.append(ZP_cd_to1)
    arr_headings_output_tregub.append('ZP_cd_to2')
    arr_value_output_tregub.append(ZP_cd_to2)
    arr_headings_output_tregub.append('ZP_cd_tr')
    arr_value_output_tregub.append(ZP_cd_tr)
    arr_headings_output_tregub.append('ZP_cd_all')
    arr_value_output_tregub.append(ZP_cd_all)
    arr_headings_output_tregub.append('P_RFZP_to1')
    arr_value_output_tregub.append(P_RFZP_to1)
    arr_headings_output_tregub.append('P_RFZP_to2')
    arr_value_output_tregub.append(P_RFZP_to2)
    arr_headings_output_tregub.append('P_RFZP_tr')
    arr_value_output_tregub.append(P_RFZP_tr)
    arr_headings_output_tregub.append('C_rr_sred')
    arr_value_output_tregub.append(C_rr_sred)
    arr_headings_output_tregub.append('K_d')
    arr_value_output_tregub.append(K_d)
    arr_headings_output_tregub.append('D_vc_to1')
    arr_value_output_tregub.append(D_vc_to1)
    arr_headings_output_tregub.append('D_vc_to2')
    arr_value_output_tregub.append(D_vc_to2)
    arr_headings_output_tregub.append('D_vc_tr')
    arr_value_output_tregub.append(D_vc_tr)
    arr_headings_output_tregub.append('OZP_to1')
    arr_value_output_tregub.append(OZP_to1)
    arr_headings_output_tregub.append('OZP_to2')
    arr_value_output_tregub.append(OZP_to2)
    arr_headings_output_tregub.append('OZP_tr')
    arr_value_output_tregub.append(OZP_tr)
    arr_headings_output_tregub.append('DZP_to1')
    arr_value_output_tregub.append(DZP_to1)
    arr_headings_output_tregub.append('DZP_to2')
    arr_value_output_tregub.append(DZP_to2)
    arr_headings_output_tregub.append('DZP_tr')
    arr_value_output_tregub.append(DZP_tr)
    arr_headings_output_tregub.append('OFZP_to1')
    arr_value_output_tregub.append(OFZP_to1)
    arr_headings_output_tregub.append('OFZP_to2')
    arr_value_output_tregub.append(OFZP_to2)
    arr_headings_output_tregub.append('OFZP_tr')
    arr_value_output_tregub.append(OFZP_tr)
    arr_headings_output_tregub.append('SOC_OTC_to1')
    arr_value_output_tregub.append(SOC_OTC_to1)
    arr_headings_output_tregub.append('SOC_OTC_to2')
    arr_value_output_tregub.append(SOC_OTC_to2)
    arr_headings_output_tregub.append('SOC_OTC_tr')
    arr_value_output_tregub.append(SOC_OTC_tr)
    arr_headings_output_tregub.append('OFZP_all_with_otc_to1')
    arr_value_output_tregub.append(OFZP_all_with_otc_to1)
    arr_headings_output_tregub.append('OFZP_all_with_otc_to2')
    arr_value_output_tregub.append(OFZP_all_with_otc_to2)
    arr_headings_output_tregub.append('OFZP_all_with_otc_tr')
    arr_value_output_tregub.append(OFZP_all_with_otc_tr)
append_2_9()

# 3 МАТЕРИАЛЬНЫЕ ЗАТРАТЫ
day_job = table_1[VARIANT + 1][9].value
S_proizv_pom = table_1[VARIANT + 1][11].value
H_mp = table_1[VARIANT + 1][18].value  # Норма расхода на нодин литр
# 3.1 Затраты на воду
P_hoz_bit = round(13.5 * N_rr * day_job, 2)
P_teh = round(H_mp * S_proizv_pom * day_job, 2)
P_teh_hoz_bit = P_hoz_bit + P_teh
C_1l = table_1[VARIANT + 1][37].value
Z_voda = round(C_1l * P_teh_hoz_bit, 2)

# 3.2. Затраты на спецодежду
C_1k = table_1[VARIANT + 1][17].value
Z_so = C_1k * N_rr

# 3.3. Затраты на освещение
N_osv = table_1[VARIANT + 1][26].value
t_osv = table_1[VARIANT + 1][19].value
C_1kVt = table_1[VARIANT + 1][38].value
Z_osv = round(((N_osv * S_proizv_pom * t_osv * day_job)/1000) * C_1kVt, 2)

# 3.4. Затраты на силовую энергию
N_k_v = table_1[VARIANT + 1][20].value
T_sil = table_1[VARIANT + 1][21].value
K_vr = table_1[VARIANT + 1][22].value
Z_sil = round(N_k_v * T_sil * K_vr * C_1kVt, 2)

# 3.5. Затраты на тепловую энергию
H_teplo = table_1[VARIANT + 1][23].value
C_1gkall = table_1[VARIANT + 1][24].value
height_proizv_pom = table_1[VARIANT + 1][27].value
V_proizv_pom = round(S_proizv_pom * height_proizv_pom, 2)
Z_teplo = round(H_teplo * C_1gkall * V_proizv_pom, 2)
Z_obh_hoz = Z_voda + Z_so + Z_osv + Z_sil + \
    Z_teplo + OFZP_rss_vr_mon + SOC_OTC_rss_vr_mon

# 3.6. Затраты на расходные материалы и запасные части для ремонтной зоны
def find_H_m(TYPE_VEHICLE, vid_to_tr):
    if TYPE_VEHICLE == 'автомобили легковые':
        return table_3[vid_to_tr+1][1].value
    elif TYPE_VEHICLE == 'автомобили грузовые':
        return table_3[vid_to_tr+1][2].value
    elif TYPE_VEHICLE == 'автобусы':
        return table_3[vid_to_tr+1][3].value
    else:
        return table_3[vid_to_tr+1][2].value
def find_H_zch(TYPE_VEHICLE):
    if TYPE_VEHICLE == 'автомобили легковые':
        return table_3[7][1].value
    elif TYPE_VEHICLE == 'автомобили грузовые':
        return table_3[7][2].value
    elif TYPE_VEHICLE == 'автобусы':
        return table_3[7][3].value
    else:
        return table_3[7][2].value

Lg_all = sum(L_g)
category = table_1[VARIANT + 1][31].value
k_1 = table_4[category+1][1].value
k_3 = find_k_3(weather)
Z_mat_to_all = []
Z_mat_tr_all = []
Z_mat_to_tr_all = []
Z_zch_all = []
for i in range(LEN_CAR):
    if cars[i].modif_vihicle == 'седельные тягачи':
        k_2 = 1.05
    else:
        k_2 = 1
    
    norma_z_mat_to1 = find_H_m(cars[i].type_vihicle, 1)
    norma_z_mat_to2 = find_H_m(cars[i].type_vihicle, 2)
    Z_mat_to = round(((norma_z_mat_to1 + norma_z_mat_to2) * L_g[i])/1000, 2)
    Z_mat_to_all.append(Z_mat_to)
    norma_z_mat_tr = find_H_m(cars[i].type_vihicle, 3)
    Z_mat_tr = round((norma_z_mat_tr * L_g[i])/1000, 2)
    Z_mat_tr_all.append(Z_mat_tr)
    Z_mat_to_tr = Z_mat_to + Z_mat_tr
    Z_mat_to_tr_all.append(Z_mat_to_tr)
    norma_z_zch = find_H_zch(cars[i].type_vihicle)
    Z_zch = round(((norma_z_zch * L_g[i])/1000)*k_1*k_2*k_3, 2)
    Z_zch_all.append(Z_zch)

    arr_headings_output_tregub.append(f'norma_z_mat_to1{i}')
    arr_value_output_tregub.append(norma_z_mat_to1)
    arr_headings_output_tregub.append(f'norma_z_mat_to2{i}')
    arr_value_output_tregub.append(norma_z_mat_to2)
    arr_headings_output_tregub.append(f'L_g{i}')
    arr_value_output_tregub.append(L_g[i])
    arr_headings_output_tregub.append(f'Z_mat_to{i}')
    arr_value_output_tregub.append(Z_mat_to)
    arr_headings_output_tregub.append(f'norma_z_mat_tr{i}')
    arr_value_output_tregub.append(norma_z_mat_tr)
    arr_headings_output_tregub.append(f'Z_mat_tr{i}')
    arr_value_output_tregub.append(Z_mat_tr)
    arr_headings_output_tregub.append(f'Z_mat_to_tr{i}')
    arr_value_output_tregub.append(Z_mat_to_tr)
    arr_headings_output_tregub.append(f'k_2{i}')
    arr_value_output_tregub.append(k_2)
    arr_headings_output_tregub.append(f'norma_z_zch{i}')
    arr_value_output_tregub.append(norma_z_zch)
    arr_headings_output_tregub.append(f'Z_zch{i}')
    arr_value_output_tregub.append(Z_zch)
Z_mat_and_zch = sum(Z_mat_to_tr_all) + sum(Z_zch_all)
Z_itogo = Z_voda + Z_so + Z_osv + Z_sil + Z_teplo + Z_mat_and_zch


def append_3():
    arr_headings_output_tregub.append('day_job')
    arr_value_output_tregub.append(day_job)
    arr_headings_output_tregub.append('P_hoz_bit')
    arr_value_output_tregub.append(P_hoz_bit)
    arr_headings_output_tregub.append('S_proizv_pom')
    arr_value_output_tregub.append(S_proizv_pom)
    arr_headings_output_tregub.append('H_mp')
    arr_value_output_tregub.append(H_mp)
    arr_headings_output_tregub.append('P_teh')
    arr_value_output_tregub.append(P_teh)
    arr_headings_output_tregub.append('P_teh_hoz_bit')
    arr_value_output_tregub.append(P_teh_hoz_bit)
    arr_headings_output_tregub.append('C_1l')
    arr_value_output_tregub.append(C_1l)
    arr_headings_output_tregub.append('Z_voda')
    arr_value_output_tregub.append(Z_voda)
    arr_headings_output_tregub.append('C_1k')
    arr_value_output_tregub.append(C_1k)
    arr_headings_output_tregub.append('Z_so')
    arr_value_output_tregub.append(Z_so)
    arr_headings_output_tregub.append('N_osv')
    arr_value_output_tregub.append(N_osv)
    arr_headings_output_tregub.append('t_osv')
    arr_value_output_tregub.append(t_osv)
    arr_headings_output_tregub.append('C_1kVt')
    arr_value_output_tregub.append(C_1kVt)
    arr_headings_output_tregub.append('t_osv')
    arr_value_output_tregub.append(t_osv)
    arr_headings_output_tregub.append('Z_osv')
    arr_value_output_tregub.append(Z_osv)
    arr_headings_output_tregub.append('N_k_v')
    arr_value_output_tregub.append(N_k_v)
    arr_headings_output_tregub.append('T_sil')
    arr_value_output_tregub.append(T_sil)
    arr_headings_output_tregub.append('K_vr')
    arr_value_output_tregub.append(K_vr)
    arr_headings_output_tregub.append('Z_sil')
    arr_value_output_tregub.append(Z_sil)
    arr_headings_output_tregub.append('H_teplo')
    arr_value_output_tregub.append(H_teplo)
    arr_headings_output_tregub.append('C_1gkall')
    arr_value_output_tregub.append(C_1gkall)
    arr_headings_output_tregub.append('height_proizv_pom')
    arr_value_output_tregub.append(height_proizv_pom)
    arr_headings_output_tregub.append('V_proizv_pom')
    arr_value_output_tregub.append(V_proizv_pom)
    arr_headings_output_tregub.append('Z_teplo')
    arr_value_output_tregub.append(Z_teplo)
    arr_headings_output_tregub.append('Z_obh_hoz')
    arr_value_output_tregub.append(Z_obh_hoz)
    arr_headings_output_tregub.append('k_1')
    arr_value_output_tregub.append(k_1)
    arr_headings_output_tregub.append('k_3')
    arr_value_output_tregub.append(k_3)
    arr_headings_output_tregub.append('Z_mat_to_all')
    arr_value_output_tregub.append(sum(Z_mat_to_all))
    arr_headings_output_tregub.append('Z_mat_tr_all')
    arr_value_output_tregub.append(sum(Z_mat_tr_all))
    arr_headings_output_tregub.append('Z_mat_to_tr_all')
    arr_value_output_tregub.append(sum(Z_mat_to_tr_all))
    arr_headings_output_tregub.append('Z_zch_all')
    arr_value_output_tregub.append(sum(Z_zch_all))
    arr_headings_output_tregub.append('Z_mat_and_zch')
    arr_value_output_tregub.append(Z_mat_and_zch)
    arr_headings_output_tregub.append('Z_itogo')
    arr_value_output_tregub.append(Z_itogo)
append_3()


# 6 СОСТАВЛЕНИЕ СМЕТЫ ЗАТРАТ И КАЛЬКУЛЯЦИИ СЕБЕСТОИМОСТИ ТО И ТР
N_to = round(table_1[VARIANT + 1][42].value + table_1[VARIANT + 1][43].value)
S_1_to = round(((OFZP_to1 + OFZP_to2 + SOC_OTC_to1 + SOC_OTC_to2 + sum(Z_mat_to_all) + (round(Z_obh_hoz * 0.3, 2))) / (table_1[VARIANT + 1][42].value + table_1[VARIANT + 1][43].value)), 2)
C_1_to = round(1.5 * S_1_to, 2)
S_1_chel_tr = round(((OFZP_tr + SOC_OTC_tr + sum(Z_mat_tr_all) + sum(Z_zch_all) +
                    (round(Z_obh_hoz * 0.7, 2))) / table_1[VARIANT + 1][1].value), 2)
C_1_chel_tr = round(1.5 * S_1_chel_tr, 2)
dohod_to = C_1_to * N_to
dohod_tr = T_tr * C_1_chel_tr
dohod_all = dohod_to + dohod_tr


def append_6():
    arr_headings_output_tregub.append('S_1_to')
    arr_value_output_tregub.append(S_1_to)
    arr_headings_output_tregub.append('C_1_to')
    arr_value_output_tregub.append(C_1_to)
    arr_headings_output_tregub.append('N_to')
    arr_value_output_tregub.append(N_to)

    arr_headings_output_tregub.append('T_tr')
    arr_value_output_tregub.append(T_tr)
    arr_headings_output_tregub.append('S_1_chel_tr')
    arr_value_output_tregub.append(S_1_chel_tr)
    arr_headings_output_tregub.append('C_1_chel_tr')
    arr_value_output_tregub.append(C_1_chel_tr)
    arr_headings_output_tregub.append('dohod_to')
    arr_value_output_tregub.append(dohod_to)
    arr_headings_output_tregub.append('dohod_tr')
    arr_value_output_tregub.append(dohod_tr)
    arr_headings_output_tregub.append('dohod_all')
    arr_value_output_tregub.append(dohod_all)

append_6()


# 7.4. Затраты на приобретение оборудования
zatrati_na_new_oborud = round(table_1[VARIANT + 1][45].value, 2)
Z_mont_dem_trans = round(zatrati_na_new_oborud * 0.25, 2)
zatrati_na_new_oborud_with_mont = zatrati_na_new_oborud + Z_mont_dem_trans
def append_7():
    arr_headings_output_tregub.append('zatrati_na_new_oborud')
    arr_value_output_tregub.append(zatrati_na_new_oborud)
    arr_headings_output_tregub.append('Z_mont_dem_trans')
    arr_value_output_tregub.append(Z_mont_dem_trans)
    arr_headings_output_tregub.append('zatrati_na_new_oborud_with_mont')
    arr_value_output_tregub.append(zatrati_na_new_oborud_with_mont)
append_7()


row_1 = dohod_all
row_2 = Z_mat_and_zch
row_3 = OFZP_to1 + OFZP_to2 + OFZP_tr
row_4 = SOC_OTC_to1 + SOC_OTC_to2 + SOC_OTC_tr
row_5 = row_2 + row_3 + row_4
row_6 = 0
row_7 = OFZP_rss_vr_mon + SOC_OTC_rss_vr_mon
row_8 = 0
row_9 = 0
row_10 = Z_voda + Z_so + Z_osv + Z_sil + Z_teplo
row_11 = 0
row_12 = row_6 + row_7 + row_8 + row_9 + row_10 + row_11
row_13 = row_5 + row_12



Zatrati_08 = row_12 + row_5
P_n =  round(dohod_all - Zatrati_08,2)
#10.2 Сумма отчислений налога на прибыль
ONPp = round(P_n * 0.2,2)
#10.3 Чистая прибыль
P_ch = round(P_n - ONPp,2)
#10.5 Определение срока окупаемости
srok_okup = round(zatrati_na_new_oborud_with_mont / P_ch , 4)

def append_8():
    arr_headings_output_tregub.append('Zatrati_08')
    arr_value_output_tregub.append(Zatrati_08)
    arr_headings_output_tregub.append('P_n')
    arr_value_output_tregub.append(P_n)
    arr_headings_output_tregub.append('ONPp')
    arr_value_output_tregub.append(ONPp)
    arr_headings_output_tregub.append('P_ch')
    arr_value_output_tregub.append(P_ch)
    arr_headings_output_tregub.append('zatrati_na_new_oborud_with_mont')
    arr_value_output_tregub.append(zatrati_na_new_oborud_with_mont)
    arr_headings_output_tregub.append('srok_okup')
    arr_value_output_tregub.append(srok_okup)

append_8()





def append_for_ishod_date():
    arr_headings_output_tregub.append('T_sm')
    arr_value_output_tregub.append(table_1[VARIANT + 1][8].value)
    arr_headings_output_tregub.append('A_m_day_job')
    arr_value_output_tregub.append(table_1[VARIANT + 1][10].value)
    arr_headings_output_tregub.append('S_ofis_pom')
    arr_value_output_tregub.append(table_1[VARIANT + 1][12].value)
    arr_headings_output_tregub.append('Kol-vo postov')
    arr_value_output_tregub.append(table_1[VARIANT + 1][25].value)
    arr_headings_output_tregub.append('day_km')
    arr_value_output_tregub.append(table_1[VARIANT + 1][29].value)
    arr_headings_output_tregub.append('a_ttt')
    arr_value_output_tregub.append(table_1[VARIANT + 1][30].value)
    arr_headings_output_tregub.append('category')
    arr_value_output_tregub.append(category)
    arr_headings_output_tregub.append('weather')
    arr_value_output_tregub.append(weather)
    arr_headings_output_tregub.append('Kol-vo a_m')
    arr_value_output_tregub.append(((table_1[VARIANT + 1][33].value).split(";"))[0])
append_for_ishod_date()
# append tables
def table_2_1(tables):
    tables.cell(row=1, column=1).value = N_to1
    tables.cell(row=2, column=1).value = N_to2
    tables.cell(row=3, column=1).value = N_tr
    tables.cell(row=4, column=1).value = N_rr

    tables.cell(row=1, column=2).value = FZP_to_1
    tables.cell(row=2, column=2).value = FZP_to_2
    tables.cell(row=3, column=2).value = FZP_tr
    tables.cell(row=4, column=2).value = FZP_all

    tables.cell(row=1, column=3).value = D_vred_to1
    tables.cell(row=2, column=3).value = D_vred_to2
    tables.cell(row=3, column=3).value = D_vred_tr
    tables.cell(row=4, column=3).value = D_vred_all

    tables.cell(row=1, column=4).value = OFZP_s_uchetom_D_to1
    tables.cell(row=2, column=4).value = OFZP_s_uchetom_D_to2
    tables.cell(row=3, column=4).value = OFZP_s_uchetom_D_tr
    tables.cell(row=4, column=4).value = OFZP_s_uchetom_D_all

    tables.cell(row=1, column=5).value = K_ud_to1
    tables.cell(row=2, column=5).value = K_ud_to2
    tables.cell(row=3, column=5).value = K_ud_tr
    tables.cell(row=4, column=5).value = 1

    tables.cell(row=1, column=6).value = C_cd_to1
    tables.cell(row=2, column=6).value = C_cd_to2
    tables.cell(row=3, column=6).value = C_cd_tr
    tables.cell(row=4, column=6).value = C_cd

    tables.cell(row=5, column=1).value = "table 2_1"
def table_2_2(tables):
    tables.cell(row=6, column=1).value = K_ud_to1
    tables.cell(row=7, column=1).value = K_ud_to2
    tables.cell(row=8, column=1).value = K_ud_tr
    tables.cell(row=9, column=1).value = 1

    tables.cell(row=6, column=2).value = N_to1
    tables.cell(row=7, column=2).value = N_to2
    tables.cell(row=8, column=2).value = N_tr
    tables.cell(row=9, column=2).value = N_rr

    tables.cell(row=6, column=3).value = ZP_cd_to1
    tables.cell(row=7, column=3).value = ZP_cd_to2
    tables.cell(row=8, column=3).value = ZP_cd_tr
    tables.cell(row=9, column=3).value = ZP_cd_all

    tables.cell(row=6, column=4).value = P_RFZP_to1
    tables.cell(row=7, column=4).value = P_RFZP_to2
    tables.cell(row=8, column=4).value = P_RFZP_tr
    tables.cell(row=9, column=4).value = P_RFZP_all

    tables.cell(row=6, column=5).value = P_RFZP_to1_ZP_cd_to1
    tables.cell(row=7, column=5).value = P_RFZP_to2_ZP_cd_to2
    tables.cell(row=8, column=5).value = P_RFZP_tr_ZP_cd_tr
    tables.cell(row=9, column=5).value = P_RFZP_all_ZP_cd_all

    tables.cell(row=6, column=6).value = round(
        P_RFZP_to1_ZP_cd_to1/(12*N_to1), 2)
    tables.cell(row=7, column=6).value = round(
        P_RFZP_to2_ZP_cd_to2/(12*N_to2), 2)
    tables.cell(row=8, column=6).value = round(P_RFZP_tr_ZP_cd_tr/(12*N_tr), 2)
    tables.cell(row=9, column=6).value = round(
        P_RFZP_all_ZP_cd_all/(12*N_rr), 2)

    tables.cell(row=10, column=1).value = "table 2_2"
def table_6_1(tables):
    tables.cell(row=11, column=1).value = OFZP_to1 + OFZP_to2
    tables.cell(row=12, column=1).value = SOC_OTC_to1 + SOC_OTC_to2
    tables.cell(row=13, column=1).value = sum(Z_mat_to_all)
    tables.cell(row=14, column=1).value = round(Z_obh_hoz * 0.3, 2)
    tables.cell(row=15, column=1).value = OFZP_to1 + OFZP_to2 + SOC_OTC_to1 + SOC_OTC_to2 + sum(Z_mat_to_all) + (round(Z_obh_hoz * 0.3, 2))

    tables.cell(row=11, column=2).value = round(((OFZP_to1 + OFZP_to2) / Lg_all) * 1000, 2)
    tables.cell(row=12, column=2).value = round(((SOC_OTC_to1 + SOC_OTC_to2) / Lg_all) * 1000, 2)
    tables.cell(row=13, column=2).value = round(((sum(Z_mat_to_all)) / Lg_all) * 1000, 2)
    tables.cell(row=14, column=2).value = round(((round(Z_obh_hoz * 0.3, 2)) / Lg_all) * 1000, 2)
    tables.cell(row=15, column=2).value = round(((OFZP_to1 + OFZP_to2 + SOC_OTC_to1 + SOC_OTC_to2 + sum(Z_mat_to_all) + (round(Z_obh_hoz * 0.3, 2))) / Lg_all) * 1000, 2)

    tables.cell(row=11, column=3).value = round(((OFZP_to1 + OFZP_to2) / N_to), 2)
    tables.cell(row=12, column=3).value = round(((SOC_OTC_to1 + SOC_OTC_to2) / N_to), 2)
    tables.cell(row=13, column=3).value = round(((sum(Z_mat_to_all)) / N_to), 2)
    tables.cell(row=14, column=3).value = round(((round(Z_obh_hoz * 0.3, 2)) / N_to), 2)
    tables.cell(row=15, column=3).value = S_1_to

    tables.cell(row=16, column=1).value = "table 6_1"
def table_6_2(tables):
    tables.cell(row=17, column=1).value = OFZP_tr
    tables.cell(row=18, column=1).value = SOC_OTC_tr
    tables.cell(row=19, column=1).value = sum(Z_mat_tr_all)
    tables.cell(row=20, column=1).value = sum(Z_zch_all)
    tables.cell(row=21, column=1).value = round(Z_obh_hoz * 0.7, 2)
    tables.cell(row=22, column=1).value = OFZP_tr + SOC_OTC_tr + sum(Z_mat_tr_all) +  sum(Z_zch_all) + (round(Z_obh_hoz * 0.7, 2))

    tables.cell(row=17, column=2).value = round((OFZP_tr / Lg_all) * 1000, 2)
    tables.cell(row=18, column=2).value = round((SOC_OTC_tr / Lg_all) * 1000, 2)
    tables.cell(row=19, column=2).value = round((sum(Z_mat_tr_all) / Lg_all) * 1000, 2)
    tables.cell(row=20, column=2).value = round((sum(Z_zch_all) / Lg_all) * 1000, 2)
    tables.cell(row=21, column=2).value = round((round(Z_obh_hoz * 0.7, 2) / Lg_all) * 1000, 2)
    tables.cell(row=22, column=2).value = round(((OFZP_tr + SOC_OTC_tr + sum(Z_mat_tr_all) + sum(Z_zch_all)+ (round(Z_obh_hoz * 0.7, 2))) / Lg_all) * 1000, 2)

    tables.cell(row=17, column=3).value = round((OFZP_tr / T_tr), 2)
    tables.cell(row=18, column=3).value = round((SOC_OTC_tr / T_tr), 2)
    tables.cell(row=19, column=3).value = round((sum(Z_mat_tr_all) / T_tr), 2)
    tables.cell(row=20, column=3).value = round((sum(Z_zch_all) / T_tr), 2)
    tables.cell(row=21, column=3).value = round((round(Z_obh_hoz * 0.7, 2) / T_tr), 2)
    tables.cell(row=22, column=3).value = S_1_chel_tr

    tables.cell(row=23, column=1).value = "table 6_2"






def append_tables():
    book = openpyxl.Workbook()
    tables = book.active
    table_2_1(tables)
    table_2_2(tables)
    table_6_1(tables)
    table_6_2(tables)


    
    book.save('tables.xlsx')
    book.close()


append_tables()
init_data.close()

book1 = openpyxl.Workbook()
data_output = book1.active
for i in range(len(arr_headings_output_tregub)):
    print(arr_headings_output_tregub[i], arr_value_output_tregub[i])
    data_output.cell(row=1, column=i+1).value = arr_headings_output_tregub[i]
    data_output.cell(row=2, column=i+1).value = arr_value_output_tregub[i]
book1.save('output_treg.xlsx')
book1.close()
