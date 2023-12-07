import os
import sys
import win32com.client
import matplotlib.pyplot as plt
from datetime import datetime
from config import pn_set_2, per_up_set, per_down_set, dp_set, Unom_set, figure_set
from config import Umax_set, Umin_set, tan_min, tan_max, pn_set_4, Umax_zd_set, Umin_zd_set
from config import rmax_set, rmin_set, xmin_set, xmax_set, gmin_set, gmax_set, bmin_set, bmax_set
from config import Pnmax_set, Pnmin_set, Qnmax_set, Qnmin_set, Pgmin_set


class CheckUp:
    def __init__(self, directory_path: str):
        # ------------------------------------------------------------------
        # Получаем текущую дату и время
        self.date = f"{datetime.now().day}_{datetime.now().month}_{datetime.now().year}  {datetime.now().hour}ч_{datetime.now().minute}мин"

        # ------------------------------------------------------------------
        # Создаём папку отчетных файлов программы если её нет
        self.path_report = f"{directory_path}\Отчетные файлы"
        self.path_figure = f"{directory_path}\Отчетные файлы\Графики"
        if not os.path.exists(self.path_report):
            os.mkdir(self.path_report)
            os.mkdir(self.path_figure)
            for i in range(1, 16):
                os.mkdir(f"{self.path_report}\Критерий №{i}")

        # ------------------------------------------------------------------
        # Работа с директориями
        # Путь к директории последнего года
        for i in range(0, -len(os.listdir(directory_path)), -1):
            self.path_last_year = os.path.join(directory_path,
                                               os.listdir(directory_path)[i - 1])
            if self.path_last_year != self.path_report:
                break

        # Открываем директорию последнего года, чтобы получить количество режимов
        self.last_year = os.listdir(self.path_last_year)
        self.length = len(self.last_year)

        self.year_mass = []
        self.modes_path = []
        # Идем по всем характерным режимам
        for i in range(self.length):
            middle_mode_path = []
            for year in os.listdir(directory_path):
                if year != os.path.basename(self.path_report):
                    if i == self.length - 1:
                        # Заносим все имена папок в массив
                        self.year_mass.append(int(year))
                    # Дает путь к каждой папке года
                    year_path = os.path.join(directory_path, year)

                    # Получаем имя характерного режима
                    name_char_mode = os.listdir(year_path)[i]
                    # Путь к режиму
                    char_mode_path = os.path.join(year_path, name_char_mode)
                    # Проверка на тип файла
                    if i > 0:
                        if check != os.path.splitext(char_mode_path)[1]:
                            print("Нельзя совмещать файлы .os и .rg2")
                            sys.exit(1)
                    check = os.path.splitext(char_mode_path)[1]
                    if check == ".rg2" or check == ".os":
                        middle_mode_path.append(char_mode_path)
            self.modes_path.append(middle_mode_path)

        # Сделаем массив названий характерных режимов
        self.names_char_modes = []
        for mode in self.modes_path:
            self.names_char_modes.append(os.path.basename(mode[0]))

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод получения всех данных Растра

    def rastr_work(self, path_mode, path_sech="No"):
        # Соединяемся с Rastr
        self.rastr = win32com.client.Dispatch("Astra.Rastr")
        # Подгружаем необходимый файл
        self.rastr.load(0, path_mode, "")
        # Получаем необходимые таблицы и параметры:
        # Узлы
        self.node = self.rastr.Tables("node")
        self.ny_node = self.node.Cols("ny")
        self.u_ras = self.node.Cols("vras")
        self.na_node = self.node.Cols("na")
        self.nsx_node = self.node.Cols("nsx")
        self.pn_node = self.node.Cols("pn")
        self.qn_node = self.node.Cols("qn")
        self.vzd = self.node.Cols("vzd")
        self.u_nom = self.node.Cols("uhom")
        self.pg_node = self.node.Cols("pg")
        self.pgmax_node = self.node.Cols("pg_max")

        # Ветви
        self.vetv = self.rastr.Tables("vetv")
        self.ip = self.vetv.Cols("ip")
        self.iq = self.vetv.Cols("iq")
        self.np = self.vetv.Cols("np")
        self.i_max = self.vetv.Cols("i_max")
        self.i_dop_r = self.vetv.Cols("i_dop_r")
        self.r = self.vetv.Cols("r")
        self.x = self.vetv.Cols("x")
        self.b = self.vetv.Cols("b")
        self.g = self.vetv.Cols("g")

        # Районы
        self.area = self.rastr.Tables("area")
        self.na_area = self.area.Cols("na")
        self.nob_area = self.area.Cols("no")
        self.pn_area = self.area.Cols("pn")
        self.dp_area = self.area.Cols("dp")
        self.pop_area = self.area.Cols("pop")
        self.name_area = self.area.Cols("name")

        # Сечения
        if os.path.splitext(path_mode)[1] == ".os":
            self.sech = self.rastr.Tables("sechen")
            self.ns = self.sech.Cols("ns")
            self.psech = self.sech.Cols("psech")
            self.pmax_sech = self.sech.Cols("pmax")

        # СХН
        self.polin = self.rastr.Tables("polin")
        self.nsx_polin = self.polin.Cols("nsx")

        # Генераторы (УР)
        self.gen = self.rastr.Tables("Generator")
        self.num_gen = self.gen.Cols("Num")
        self.num_PQ_gen = self.gen.Cols("NumPQ")
        self.P_gen = self.gen.Cols("P")
        self.sta_gen = self.gen.Cols("sta")

        # PQ-диаграммы
        self.PQ = self.rastr.Tables("graphik2")
        self.num_PQ = self.PQ.Cols("Num")

        self.rastr.rgm('')

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод получения всех районов, где Pнаг > pn_set_2 (заданного)

    def give_na(self, pn=pn_set_2):
        self.rastr_work(os.path.join(self.path_last_year, self.last_year[0]))
        self.area.SetSel(f"pn>{pn}")
        na = []
        index = self.area.FindNextSel(-1)
        na.append(self.na_area.Z(index))
        for j in range(1, self.area.Count):
            index = self.area.FindNextSel(index)
            na.append(self.na_area.Z(index))
        return na

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для построения графиков

    def plot_figure(self, name, x, y, char_mode_name):
        plt.title(f"Изменение Pпотр района: {name}")
        plt.xlabel('Года', fontsize=12, color='blue')
        plt.ylabel('Pпотр', fontsize=12, color='blue')
        plt.grid()
        plt.plot(x, y, label=char_mode_name,
                 marker='o', markersize=5)
        plt.legend(fontsize=10, bbox_to_anchor=(1, 0.5))

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 1-го критерия по потерям

    def crit_1(self):
        print("Выполняется критерий №1")
        na_str = input(
            "Введите номера интересующих районов через запятую, либо нажмите 'Enter' для проверки всeх:")
        print("Начинается выполнение, пожалуйста подождите")
        if na_str != "":
            na = [int(x) for x in na_str.split(",")]
        else:
            na = self.give_na(0)
        flag = False
        # Перебор всех характерных режимов по всем годам
        for mode in self.modes_path:
            # Название характерного режима
            name_char_mode = os.path.basename(mode[0])[:-9]
            k = 0
            for mode_year in mode:
                # Подгрузка используемого режима
                self.rastr_work(mode_year)
                # Перебор всех районов
                for n in na:
                    # Выборка и получение значений
                    self.area.SetSel(f"na={n}")
                    ind = self.area.FindNextSel(-1)
                    nm_ar = self.name_area.Z(ind)
                    # Процент потерь от потребления
                    percent = self.dp_area.Z(ind) * 100 / self.pop_area.Z(ind)
                    # Проверка значения потерь (Dp) 
                    if percent > dp_set:
                        text = (f"""Warning: превышение заданного ({dp_set}%-го) значения потерь от потребления
                                Район: {nm_ar}
                                Характерный режим: {name_char_mode}
                                Год: {self.year_mass[k]}
                                Значение потерь: {round(self.dp_area.Z(ind), 2)} МВт
                                Значение потребления: {round(self.pop_area.Z(ind), 2)} МВт
                                Потери от потребления: {round(percent, 2)}%
                                """)
                        with open(f"{self.path_report}\Критерий №1\К1  {self.date}.txt", mode="a+") as f:
                            f.write(f"\n{text}")
                        flag = True
                k += 1
        if flag != True:
            text = f"Значения потерь в районах корректны"
            with open(f"{self.path_report}\Критерий №1\К1  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 2-го критерия по потреблению

    def crit_2(self, figure=figure_set):
        print("Выполняется критерий №2")
        na_str = input(
            "Введите номера интересующих районов через запятую, либо нажмите 'Enter' для проверки всeх:")
        print("Начинается выполнение, пожалуйста подождите")
        if na_str != "":
            na = [int(x) for x in na_str.split(",")]
        else:
            na = self.give_na()
        # Создаем папку для графиков
        os.mkdir(f"{self.path_figure}\\{self.date}")
        # Получим значения Pпотр для всех районов и всех режимов за все года
        flag2 = False
        # Перебор всех характерных режимов по всем годам
        full_Pop_mass_mode = []
        name_char_mass = []
        for mode in self.modes_path:
            Pop_mass_mode = []
            # Название характерного режима
            name_char_mode = os.path.basename(mode[0])[:-9]
            name_char_mass.append(name_char_mode)
            for mode_year in mode:
                # Подгрузка используемого режима
                self.rastr_work(mode_year)
                # Перебор районов
                i = 0
                nm_ar_mass = []
                for n in na:
                    # Выборка и получение значений
                    self.area.SetSel(f"na={n}")
                    ind = self.area.FindNextSel(-1)
                    nm_ar = self.name_area.Z(ind)
                    nm_ar_mass.append(nm_ar)
                    if mode_year == mode[0]:
                        Pop_mass_mode.append([round(self.pop_area.Z(ind),2)])
                    else:
                        Pop_mass_mode[i].append(round(self.pop_area.Z(ind),2))
                    i += 1
            full_Pop_mass_mode.append(Pop_mass_mode)
            
            # Проверка в соответсвии с заданными значениями per_up_set и per_down_set
            for pop in Pop_mass_mode:
                k = 0  # Предыдущее значение P
                j = 0  # Для счетчика
                for p in pop:
                    flag = False
                    if k != 0:
                        year_1 = self.year_mass[j - 1]
                        year_2 = self.year_mass[j]
                        if round((k - p) * 100 / k, 2) > per_down_set:
                            text_reason = f"снижение Pпотр на {round((k - p) * 100 / k, 2)}% ({round(k - p, 1)} МВт)"
                            flag = True
                        if p > (1 + per_up_set / 100) * k * (year_2 - year_1):
                            text_reason = f"увеличение Pпотр на {round((p - k) * 100 / k, 2)}% ({round(p - k, 1)} МВт)"
                            flag = True
                        if flag:
                            text = (f"""Warning: {text_reason}
                                Район: {nm_ar_mass[Pop_mass_mode.index(pop)]}
                                Характерный режим: {name_char_mode}
                                Года:
                                        {year_1} - Pпотр = {k}
                                        {year_2} - Pпотр = {p}
                                """)
                            with open(f"{self.path_report}\Критерий №2\К2  {self.date}.txt", mode="a+") as f:
                                f.write(f"\n{text}")
                            flag2 = True
                    k = p
                    j += 1
        # Строим график если нужно            
        for n in na:
            for fpop in full_Pop_mass_mode:
                if figure:
                    index = full_Pop_mass_mode.index(fpop)
                    nm_char = name_char_mass[index]
                    if nm_char.find("Зима") != -1:
                        label = nm_char[nm_char.find("Зима"):]
                    elif nm_char.find("Лето") != -1:
                        label = nm_char[nm_char.find("Лето"):]
                    else:
                        label = nm_char[nm_char.find("Паводок"):]
                    self.plot_figure(nm_ar_mass[na.index(n)], self.year_mass, fpop[na.index(n)], label)
            if figure:
                plt.savefig(f'{self.path_figure}\\{self.date}\\{nm_ar_mass[na.index(n)]}.png',
                            bbox_inches='tight')
                plt.close()
            
        if flag2 != True:
            text = f"P_потр монотонно изменяется в заданных пределах"
            with open(f"{self.path_report}\Критерий №2\К2  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 3-го критерия по расчетному напряжению

    def crit_3(self):
        print("Выполняется критерий №3")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                k = 0
                # Перебор всех Uном
                for Unom in Unom_set:
                    self.node.SetSel(f"uhom={Unom}")
                    index = self.node.FindNextSel(-1)
                    # Перебор всех узлов с рассматриваемым Uном
                    mass_ny = []
                    for i in range(self.node.Count):
                        # Проверка условия
                        u_ras = self.u_ras.Z(index)
                        if u_ras > 0 and (u_ras < (Umin_set / 100) * Unom or u_ras > Umax_set[k]):
                            mass_ny.append([self.ny_node.Z(index),round(u_ras,2)])
                        index = self.node.FindNextSel(index)
                    if len(mass_ny) > 0:
                        text = (f"""Warning: расчетные значения напряжений выходят из заданных пределов
                                    Год: {self.year_mass[j]}
                                    Характерный режим: {name_char_mode}
                                    Uном: {Unom} кВ
                                    Номера узлов и Uрасч в кВ: {mass_ny}
                                    """)
                        with open(f"{self.path_report}\Критерий №3\К3  {self.date}.txt", mode="a+") as f:
                            f.write(f"\n{text}")
                        flag = True
                    k += 1
                j += 1
        if flag != True:
            text = f"Расчетные значения напряжений не выходят из заданных пределов"
            with open(f"{self.path_report}\Критерий №3\К3  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 4-го критерия по отклонению U_зд от номинального

    def crit_4(self):
        print("Выполняется критерий №4")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                self.node.SetSel("vzd>0")
                index = self.node.FindNextSel(-1)
                # Перебор всех узлов
                mass_ny = []
                for i in range(self.node.Count):
                    u_zd = self.vzd.Z(index)
                    u_nom = self.u_nom.Z(index)
                    # Проверка условия
                    if u_zd > u_nom * (1 + Umax_zd_set / 100) or u_zd < u_nom * (1 - Umin_zd_set / 100):
                        mass_ny.append([self.ny_node.Z(index),round(u_zd,2)])
                    index = self.node.FindNextSel(index)
                if len(mass_ny) > 0:
                    text = (
                        f"""Warning: заданное напряжение (U_зд) выходит из заданных пределов (-{Umin_zd_set}%; +{Umax_zd_set}%)
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов и Uзд в кВ: {mass_ny}
                                """)
                    with open(f"{self.path_report}\Критерий №4\К4  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все заданные напряжения имеют корректные значения"
            with open(f"{self.path_report}\Критерий №4\К4  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 5-го критерия по соответсвтвию P и Q нагрузки

    def crit_5(self):
        print("Выполняется критерий №5")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                # Делаем выборку
                self.node.SetSel(f"pn>={pn_set_4}")
                index = self.node.FindNextSel(-1)
                # Перебор всех узлов с P_нагн >= pn_set_4
                for i in range(self.node.Count):
                    # Проверка условия
                    pn = self.pn_node.Z(index)
                    qn = self.qn_node.Z(index)
                    tan = qn / pn
                    if tan < tan_min or tan > tan_max:
                        text = (f"""Warning: отношение Q_нагр к P_нагр выходит из заданных пределов
                            Год: {self.year_mass[j]}
                            Характерный режим: {name_char_mode}
                            Номер узла: {self.ny_node.Z(index)}
                            Q_нагр: {round(qn,2)}
                            P_нагр: {round(pn,2)}
                            Тангенс: {round(tan,2)}
                            """)
                        with open(f"{self.path_report}\Критерий №5\К5  {self.date}.txt", mode="a+") as f:
                            f.write(f"\n{text}")
                        flag = True
                    index = self.node.FindNextSel(index)
                j += 1
        if flag != True:
            text = f"Все узлы имеют корректное отношение Q_нагр к P_нагр"
            with open(f"{self.path_report}\Критерий №5\К5  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 6-го критерия по привязке узлов к районам

    def crit_6(self):
        print("Выполняется критерий №6")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                index = self.node.FindNextSel(-1)
                # Перебор всех узлов
                mass_ny = []
                for i in range(self.node.Count):
                    # Проверка условия
                    if self.na_node.Z(index) < 1:
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.node.FindNextSel(index)
                if len(mass_ny) > 0:
                    text = (f"""Warning: отсутствует привязка узлов к районам
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\Критерий №6\К6  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все узлы привязаны к районам"
            with open(f"{self.path_report}\Критерий №6\К6  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 7-го критерия по заданию номера объединения в районах
    def crit_7(self):
        print("Выполняется критерий №7")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                index = self.area.FindNextSel(-1)
                # Перебор всех районов
                mass_na = []
                for i in range(self.area.Count):
                    # Проверка условия
                    if self.nob_area.Z(index) < 1:
                        mass_na.append(self.na_area.Z(index))
                    index = self.area.FindNextSel(index)
                if len(mass_na) > 0:
                    text = (f"""Warning: отсутствуют номера объединений районов
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера районов: {mass_na}
                                """)
                    with open(f"{self.path_report}\Критерий №7\К7  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все районы имеют заданный номер объединения"
            with open(f"{self.path_report}\Критерий №7\К7  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 8-го критерия по превышению перетоков в сечениях

    def crit_8(self):
        if os.path.splitext(self.modes_path[0][0])[1]!=".os":
            print("Критерий №8 НЕ выполнен. Формат считываемых файлов не '.os'")
            return
        print("Выполняется критерий №8")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                path_sech = mode_year
                self.rastr_work(mode_year, path_sech)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-3]
                self.sech.SetSel("pmax>0")
                index = self.sech.FindNextSel(-1)
                # Перебор всех сечений
                mass_ns = []
                for i in range(self.sech.Count):
                    val_sch = self.psech.Z(index)
                    max_sch = self.pmax_sech.Z(index)
                    # Проверка условия
                    if val_sch > max_sch:
                        mass_ns.append([self.ns.Z(index),round(val_sch,2),f"({(val_sch/max_sch-1)*100}%)"])
                    index = self.sech.FindNextSel(index)
                if len(mass_ns) > 0:
                    text = (f"""Warning: превышение перетоков в сечениях
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера сечений и переток по ним в МВт: {mass_ns} 
                                """)
                    with open(f"{self.path_report}\Критерий №8\К8  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все сечения имеют корректные перетоки"
            with open(f"{self.path_report}\Критерий №8\К8  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 9-го критерия по неизменности количества строк в таблицах
    def crit_9(self):
        print("Выполняется критерий №9")
        for i in range(len(self.year_mass)):
            with open(f"{self.path_report}\Критерий №9\К9  {self.date}.txt", mode="a+") as f:
                f.write(f"\n\n{self.year_mass[i]} год\n[Узлы, Ветви, Районы, Терр, Ген]")
            for k in range(len(self.modes_path)):
                # Соединяемся с Rastr
                self.rastr = win32com.client.Dispatch("Astra.Rastr")
                # Подгружаем необходимый файл
                self.rastr.load(0, self.modes_path[k][i], "")
                amount_node = self.rastr.Tables("node").Count
                amount_vetv = self.rastr.Tables("vetv").Count
                amount_area = self.rastr.Tables("area").Count
                amount_area2 = self.rastr.Tables("area2").Count
                amount_Generator = self.rastr.Tables("Generator").Count
                amount_mass = [amount_node, amount_vetv, amount_area, amount_area2, amount_Generator]
                with open(f"{self.path_report}\Критерий №9\К9  {self.date}.txt", mode="a+") as f:
                    f.write(f"\n{amount_mass} - {os.path.basename(self.modes_path[k][i])}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 10-го критерия по соответствию используемых СХН
    def crit_10(self):
        print("Выполняется критерий №10")
        flag2 = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                # Номера имеющихся СХН
                mass_nsx = []
                for k in range(self.polin.Count):
                    mass_nsx.append(self.nsx_polin.Z(k))
                # Задаём массив итоговых узлов удовлетворяющих условию
                mass_ny = []
                # Задаём множество используемых СХН
                mass_nsx_use = set()
                # Выборка
                self.node.SetSel("nsx>=1")
                index = self.node.FindNextSel(-1)
                # Перебор всех узлов
                for i in range(self.node.Count):
                    flag = True
                    nsxnode = self.nsx_node.Z(index)
                    # Проверка условий
                    for nsx in mass_nsx:
                        # Проверка условие ссылки на несуществующий СХН
                        if nsxnode == nsx or nsxnode == 1 or nsxnode == 2:
                            flag = False
                        # Формируем неповторяющееся множество используемых СХН
                        if nsx == nsxnode:
                            mass_nsx_use.add(nsx)
                    if flag:
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.node.FindNextSel(index)
                # Получаем номера СХН, которые не используются
                mass_nsx = set(mass_nsx)
                mass_nsx.difference_update(mass_nsx_use)
                if len(mass_ny) > 0:
                    text = (f"""Warning: обращение к несуществующему СХН
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\Критерий №10\К10  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag2 = True
                if len(mass_nsx_use) > 0:
                    text = (f"""Warning: неиспользуемый СХН
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера CХН: {mass_nsx}
                                """)
                    with open(f"{self.path_report}\Критерий №10\К10  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag2 = True
                j += 1
        if flag2 != True:
            text = f"СХН заданны верно"
            with open(f"{self.path_report}\Критерий №10\К10  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 11-го критерия по соответствию используемых PQ-диаграмм в таблице Генераторы (УР)
    def crit_11(self):
        print("Выполняется критерий №11")
        flag2 = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                # Множество номеров имеющихся PQ-диаграмм
                mass_PQ = set()
                for k in range(self.PQ.Count):
                    mass_PQ.add(self.num_PQ.Z(k))
                # Задаём массив итоговых номеров генераторов удовлетворяющих условию
                mass_num_gen = []
                # Задаём множество используемых PQ-диаграмм
                mass_PQ_use = set()
                # Выборка
                self.gen.SetSel("NumPQ>=1")
                index = self.gen.FindNextSel(-1)
                for i in range(self.gen.Count):
                    flag = True
                    PQ_gen = self.num_PQ_gen.Z(index)
                    # Проверка условий
                    for PQ in mass_PQ:
                        # Проверка условия ссылки на несуществующую PQ-диаграмму
                        if PQ_gen == PQ:
                            flag = False
                            mass_PQ_use.add(PQ)
                    if flag:
                        mass_num_gen.append(self.num_gen.Z(index))
                    index = self.gen.FindNextSel(index)
                # Получаем номера PQ, которые не используются
                mass_PQ.difference_update(mass_PQ_use)
                if len(mass_num_gen) > 0:
                    text = (f"""Warning: обращение к несуществующей PQ-диаграмме
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера генераторов: {mass_num_gen}
                                """)
                    with open(f"{self.path_report}\Критерий №11\К11  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag2 = True
                if len(mass_PQ) > 0:
                    text = (f"""Warning: неиспользуемая PQ-диаграмма
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера PQ+диаграмм: {mass_PQ}
                                """)
                    with open(f"{self.path_report}\Критерий №11\К11  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag2 = True
                j += 1
        if flag2 != True:
            text = f"PQ-диаграммы заданны верно"
            with open(f"{self.path_report}\Критерий №11\К11  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 12-го критерия по соответствию состояния и мощности генератора
    def crit_12(self):
        print("Выполняется критерий №12")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                index = self.gen.FindNextSel(-1)
                # Перебор всех генераторов
                mass_num_gen = []
                for i in range(self.gen.Count):
                    # Проверка условия (sta=False - Значит включенный ген; sta=True - Значит отключенный ген)
                    if (self.sta_gen.Z(index) == True and self.P_gen.Z(index) != 0) or (
                            self.sta_gen.Z(index) == False and self.P_gen.Z(index) == 0):
                        mass_num_gen.append(self.num_gen.Z(index))
                    index = self.gen.FindNextSel(index)
                if len(mass_num_gen) > 0:
                    text = (f"""Warning: несоответсвие состояния и мощности генератора
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера генераторов: {mass_num_gen}
                                """)
                    with open(f"{self.path_report}\Критерий №12\К12  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все генераторы заданны верно"
            with open(f"{self.path_report}\Критерий №12\К12  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 13-го критерия по нарушению токовых ограничений
    def crit_13(self):
        print("Выполняется критерий №13")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                self.vetv.SetSel("i_dop_r>0")
                index = self.vetv.FindNextSel(-1)
                # Перебор всех ветвей
                mass_ip_iq = []
                for i in range(self.vetv.Count):
                    # Проверка нарушения токовых ограничения
                    if self.i_max.Z(index) * 1000 > self.i_dop_r.Z(index):
                        mass_ip_iq.append([self.ip.Z(index), self.iq.Z(index), self.np.Z(index)])
                    index = self.vetv.FindNextSel(index)
                if len(mass_ip_iq) > 0:
                    text = (f"""Warning: нарушение токовых ограничений
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Линии формата [нач., кон., парл.]: {mass_ip_iq}
                                """)
                    with open(f"{self.path_report}\Критерий №13\К13  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Токовые ограничения не нарушаются"
            with open(f"{self.path_report}\Критерий №13\К13  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 14-го критерия по корректности параметров узлов/ветвей 
    def crit_14(self):
        print("Выполняется критерий №14")
        flag_vetv = False
        flag_node = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                index = self.vetv.FindNextSel(-1)

                # Перебор всех ветвей
                mass_ip_iq = []
                param_vetv = set()
                for i in range(self.vetv.Count):
                    flag2 = False
                    r = self.r.Z(index)
                    x = self.x.Z(index)
                    g = self.g.Z(index) * 1e+6
                    b = self.b.Z(index) * 1e+6
                    # Проверка условий
                    if r > rmax_set or r < rmin_set:
                        param_vetv.add("R")
                        flag2 = True
                    if x > xmax_set or x < xmin_set:
                        param_vetv.add("X")
                        flag2 = True
                    if g > gmax_set or g < gmin_set:
                        param_vetv.add("G")
                        flag2 = True
                    if b > bmax_set or b < bmin_set:
                        param_vetv.add("B")
                        flag2 = True
                    if flag2:
                        mass_ip_iq.append([self.ip.Z(index), self.iq.Z(index), self.np.Z(index)])
                    index = self.vetv.FindNextSel(index)
                index = self.vetv.FindNextSel(-1)

                # Перебор всех узлов
                index = self.node.FindNextSel(-1)
                mass_ny = []
                param_node = set()
                for i in range(self.node.Count):
                    flag2 = False
                    pn = self.pn_node.Z(index)
                    qn = self.qn_node.Z(index)
                    pg = self.pg_node.Z(index)
                    pgmax = self.pgmax_node.Z(index)
                    # Проверка условий
                    if pn > Pnmax_set or pn < Pnmin_set:
                        param_node.add("Pn")
                        flag2 = True
                    if qn > Qnmax_set or qn < Qnmin_set:
                        param_node.add("Qn")
                        flag2 = True
                    if pg > pgmax or pg < Pgmin_set:
                        param_node.add("Pg")
                        flag2 = True
                    if flag2:
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.node.FindNextSel(index)

                # Запись в отчет
                if len(mass_ip_iq) > 0:
                    text = (f"""Warning: нехарактерные значения параметров ветвей
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Параметры: {param_vetv}
                                Номера ветвей вида [нач., кон., парл.]: {mass_ip_iq}
                                """)
                    with open(f"{self.path_report}\Критерий №14\К14 ВЕТВИ  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag_vetv = True
                if len(mass_ny) > 0:
                    text = (f"""Warning: нехарактерные значения параметров узлов
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Параметры: {param_node}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\Критерий №14\К14 УЗЛЫ  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag_node = True
                j += 1

        if flag_vetv != True:
            text = f"Все параметры ветвей заданны верно"
            with open(f"{self.path_report}\Критерий №14\К14 ВЕТВИ  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")
        if flag_node != True:
            text = f"Все параметры узлов заданны верно"
            with open(f"{self.path_report}\Критерий №14\К14 УЗЛЫ  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 15-го критерия по наличию СХН в узле с ненулевой нагрузкой
    def crit_15(self):
        print("Выполняется критерий №15")
        flag = False
        for mode in self.modes_path:
            j = 0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                self.node.SetSel("pn>0")
                index = self.node.FindNextSel(-1)
                # Перебор всех узлов
                mass_ny = []
                for i in range(self.node.Count):
                    # Проверка задания СХН
                    if self.nsx_node.Z(index) == 0:
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.node.FindNextSel(index)
                if len(mass_ny) > 0:
                    text = (f"""Warning: Отсутсвует СХН в узле с ненулевой нагрузкой
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номер узла: {mass_ny}
                                """)
                    with open(f"{self.path_report}\Критерий №15\К15  {self.date}.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag != True:
            text = f"Все СХН заданны"
            with open(f"{self.path_report}\Критерий №15\К15  {self.date}.txt", mode="a+") as f:
                f.write(f"\n{text}")

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки ВСЕХ критериев
    def all_crit(self):
        self.crit_1()
        self.crit_2()
        self.crit_3()
        self.crit_4()
        self.crit_5()
        self.crit_6()
        self.crit_7()
        self.crit_8()
        self.crit_9()
        self.crit_10()
        self.crit_11()
        self.crit_12()
        self.crit_13()
        self.crit_14()
        self.crit_15()


def start(direct_path):
    direct_path = input("Введите путь к папкам годов:\n")
    direct_path = direct_path.replace('"','')
    print("""Номера критериев:
    0  -  Проверка ВСЕХ критериев
    1  -  Проверка корректности заданных потерь Района;
    2  -  Проверка немонотонности изменения потребления Района;
    3  -  Проверка отклонения расчетного напряжения;
    4  -  Проверка отклонения U_зд от номинального;
    5  -  Проверка соотношения P и Q нагрузки;
    6  -  Проверка привязки узлов к районам;
    7  -  Проверка задания номера объединения в таблице "Районы";
    8  -  Проверка превышения перетоков в сечениях;
    9  -  Проверка неизменности количества строк в характерных режимах;
    10 -  Проверка взаимосвязи заданных СХН в узлах с таблицей СХН
    11 -  Проверка взаимосвязи заданных PQ-диаграмм таблицы Генераторы (УР)
    12 -  Проверка соответствия состояния и мощности генератора
    13 -  Проверка нарушения токовых ограничений по ветвям
    14 -  Проверка корректности параметров ветвей и узлов
    15 -  Проверка наличия СХН в узле ненулевой нагрузкой
          """)
    num_crit = input("Введите номер критерия: ")
    if num_crit == "0":
        CheckUp(direct_path).all_crit()
    elif num_crit == "1":
        CheckUp(direct_path).crit_1()
    elif num_crit == "2":
        CheckUp(direct_path).crit_2()
    elif num_crit == "3":
        CheckUp(direct_path).crit_3()
    elif num_crit == "4":
        CheckUp(direct_path).crit_4()
    elif num_crit == "5":
        CheckUp(direct_path).crit_5()
    elif num_crit == "6":
        CheckUp(direct_path).crit_6()
    elif num_crit == "7":
        CheckUp(direct_path).crit_7()
    elif num_crit == "8":
        CheckUp(direct_path).crit_8()
    elif num_crit == "9":
        CheckUp(direct_path).crit_9()
    elif num_crit == "10":
        CheckUp(direct_path).crit_10()
    elif num_crit == "11":
        CheckUp(direct_path).crit_11()
    elif num_crit == "12":
        CheckUp(direct_path).crit_12()
    elif num_crit == "13":
        CheckUp(direct_path).crit_13()
    elif num_crit == "14":
        CheckUp(direct_path).crit_14()
    elif num_crit == "15":
        CheckUp(direct_path).crit_15()
    else:
        print("Неверный критерий")
    return direct_path


path = ''
while True:
    flag = input("Чтобы начать новый расчет нажмите 'Enter'")
    if flag == "":
        path = start(path)
    else:
        break
