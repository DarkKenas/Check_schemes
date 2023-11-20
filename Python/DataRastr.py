from config import pn, per_up, per_down

# Фукнция получения данных по району


def DataRastr(directory, pn, per_up, per_down, na, figure, cl_rep=False, cl_fig=False):
    print("Начинается выполнение, пожалуйста подождите")
    import os
    import matplotlib.pyplot as plt

    # Создаём папку отчетных файлов программы если её нет
    path_rep = f"../Отчетные файлы"
    path_figure = f"../Отчетные файлы/Графики"
    if os.path.exists(path_rep) == False:
        os.mkdir(path_rep)
        os.mkdir(path_figure)
    else:
        if cl_rep == True:
            with open(f"{path_rep}\\report_11.txt", mode="w") as f:
                f.write("Здесь формируется отчет\n\n")
        if cl_fig == True:
            for f in os.listdir("../Отчетные файлы/Графики"):
                os.remove(os.path.join("../Отчетные файлы/Графики", f))

    # Соединяемся с Rastr
    import win32com.client
    Rastr = win32com.client.Dispatch("Astra.Rastr")

    # Путь к директории последнего года
    path_yr = os.path.join(directory, os.listdir(directory)[-1])
    # Открываем директорию последнего года
    last_yr = os.listdir(os.path.join(directory, os.listdir(directory)[-1]))

    # Если не заданы интересующие районы, то задаём все, где Pn>pn
    if na == False:
        Rastr.Load(0, os.path.join(path_yr, last_yr[0]), "")

        # Получаем номера всех районов
        area = Rastr.Tables("area")
        na_area = area.Cols("na")

        area.SetSel(f"pn>{pn}")
        na = []
        index = area.FindNextSel(-1)
        na.append(na_area.Z(index))
        for j in range(1, area.Count):
            index = area.FindNextSel(index)
            na.append(na_area.Z(index))
    Pop_mass = []  # Массив значений Pпотр района

    # Находим количество хар. режимов
    lenght = len(last_yr)
    plt.figure(figsize=(8, 7.5))
    # Перебор всех заданных районов
    for n in na:
        print(f"Проверяю район номер {n}")
        # Перебор по всем хар. режимам
        for i in range(0, lenght):
            Pop_mass = []   # Массив значений Pпотр района
            year_mass = []  # Массив годов

            # Идём по всем годам
            for year in os.listdir(directory):
                year_mass.append(int(year))
                yp = os.path.join(directory, year)    # Дает путь к папке года

                name_chm = os.listdir(yp)[i]
                # Путь к k-ому хар. режиму из папки года
                path = os.path.join(yp, name_chm)

                # Подгрузка используемых файлов
                Rastr.Load(0, path, "")

                # Получаем данные по районам
                area = Rastr.Tables("area")
                name_area = area.Cols("name")
                pop_area = area.Cols("pop")

                area.SetSel(f"na={n}")
                ind = area.FindNextSel(-1)
                pop = pop_area.Z(ind)
                Pop_mass.append(pop)
            nm_ar = name_area.Z(ind)

            l = 0
            for p in Pop_mass:
                if l != 0:
                    y1 = year_mass[Pop_mass.index(l)]
                    y2 = year_mass[Pop_mass.index(p)]
                if p < l:
                    # Изменение в процентах
                    per = round((l-p)*100/l, 2)
                    if per > per_down:
                        text_down = (f"""Warning: снижение Pпотр на {per}% ({round(l-p, 1)} МВт)
                            Район: {nm_ar}
                            Хар. режим: {name_chm}
                            Года:
                                    {y1} - Pпотр = {l}
                                    {y2} - Pпотр = {p}
                            """)
                        with open(f"{path_rep}\\report_11.txt", mode="a+") as f:
                            f.write(f"\n{text_down}")
                if l != 0 and p > (1+per_up/100)*l*(y2-y1):
                    text_up = (f"""Warning: увеличение Pпотр свыше {per_up*(y2-y1)}%
                          Район: {nm_ar}
                          Хар. режим: {name_chm}
                          Года:
                                {year_mass[Pop_mass.index(l)]}
                                {year_mass[Pop_mass.index(p)]}
                          """)
                    with open(f"{path_rep}\\report_11.txt", mode="a+") as f:
                        f.write(f"\n{text_up}")
                l = p

            # Строим график если нужно
            if figure != False:
                if name_chm.find("Зима") != -1:
                    lbl = name_chm[name_chm.find("Зима"):-9]
                elif name_chm.find("Лето") != -1:
                    lbl = name_chm[name_chm.find("Лето"):-9]
                else:
                    lbl = name_chm[name_chm.find("Паводок"):-9]
                plt.title(f"Изменение Pпотр района: {nm_ar}")
                plt.xlabel('Года', fontsize=12, color='blue')
                plt.ylabel('Pпотр', fontsize=12, color='blue')
                plt.grid()
                plt.plot(year_mass, Pop_mass, label=lbl,
                         marker='o', markersize=5)
                plt.legend(fontsize=10, bbox_to_anchor=(1, 0.5))

        if figure != False:
            print("Строю график")
            plt.savefig(f'{path_figure}\\{nm_ar}.png', bbox_inches='tight')
            plt.close()


def start():
    path_mode = input("Задайте путь к папке с годами расчетных схем:\n")
    check_na = input("""\nПроверить все районы?
                        Если да, нажать: Enter
                        Если нет, написать номера районов через запятую без пробелов: """)
    check_fig = input("""\nПостроить графики изменения Pпотр:
                        Если да, нажать: Enter
                        Если нет, написать (No): """)
    if check_na != "":
        na_mass = check_na.split(",")
        na = [int(i) for i in na_mass]
    else:
        na = False
    if check_fig == "":
        figure = True
    else:
        figure = False
    check_rep = input("""Очистить файл отчёта?
                        Если да, написать (Yes), если нет, написать (No): """)
    if check_rep == "Yes":
        cl_rep = True
    check_fig_f = input("""Очистить папку графиков?
                        Если да, написать (Yes), если нет, написать (No): """)
    if check_fig_f == "Yes":
        cl_fig = True
    # Вызов функции
    DataRastr(path_mode, pn, per_up, per_down, na, figure, cl_rep, cl_fig)

    print("Проверка завершена")


start()

while True:
    flag = input("Чтобы начать новый расчет нажмите: Enter")
    if flag == "":
        start()
    else:
        break
