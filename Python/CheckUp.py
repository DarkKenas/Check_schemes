import os
import win32com.client
import matplotlib.pyplot as plt
from config import pn_set_2, per_up_set, per_down_set, dp_set, Unom_set
from config import Umax_set, Umin_set, tan_min, tan_max, pn_set_4, Umax_zd_set, Umin_zd_set


class CheckUp:
    def __init__(self, directory_path: str, clear_report: bool, clear_figure: bool):
        # ------------------------------------------------------------------
        # Создаём папку отчетных файлов программы если её нет
        self.path_report = f"Отчетные файлы"
        self.path_figure = f"Отчетные файлы/Графики"
        if not os.path.exists(self.path_report):
            os.mkdir(self.path_report)
            os.mkdir(self.path_figure)
        else:
            # ------------------------------------------------------------------
            # Предварительная чистка отчетных файлов
            if clear_report:
                with open(f"{self.path_report}\\report_P_потери.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_P_потр.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_U_расч.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_U_зд.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_P_Q_нагр.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_узлы_к_районам.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_объединения_районов.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_перетоки_сечений.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
                with open(f"{self.path_report}\\report_количество_строк.txt", mode="w") as f:
                    f.write("Здесь формируется отчет\n\n")
            if clear_figure:
                for f in os.listdir(self.path_figure):
                    os.remove(os.path.join(self.path_figure, f))

        # ------------------------------------------------------------------
        # Работа с директориями
        # Путь к директории последнего года
        self.path_last_year = os.path.join(directory_path,
                                           os.listdir(directory_path)[-1])

        # Открываем директорию последнего года, чтобы получить количество режимов
        self.last_year = os.listdir(self.path_last_year)
        self.length = len(self.last_year)

        self.year_mass = []
        self.modes_path = []
        # Идем по всем характерным режимам
        for i in range(self.length):
            middle_mode_path = []
            for year in os.listdir(directory_path):
                if i==self.length-1:
                    # Заносим все имена папок в массив
                    self.year_mass.append(int(year))
                # Дает путь к каждой папке года
                year_path = os.path.join(directory_path, year)

                # Получаем имя характерного режима
                name_char_mode = os.listdir(year_path)[i]
                # Путь к режиму
                char_mode_path = os.path.join(year_path, name_char_mode)
                middle_mode_path.append(char_mode_path)
            self.modes_path.append(middle_mode_path)
        
        # Сделаем массив названий характерных режимов
        self.names_char_modes = []
        for mode in self.modes_path:
            self.names_char_modes.append(os.path.basename(mode[0]))

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод получения всех данных Растра

    def rastr_work(self, path_mode, path_sech="No", path_shabl_sech="No"):
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

        # Ветви
        self.vetv = self.rastr.Tables("vetv")
        self.ip = self.vetv.Cols("ip")
        self.iq = self.vetv.Cols("iq")
        self.r = self.vetv.Cols("r")
        self.x = self.vetv.Cols("x")
        self.b = self.vetv.Cols("b")
        self.na_vetv = self.vetv.Cols("na")

        # Районы
        self.area = self.rastr.Tables("area")
        self.na_area = self.area.Cols("na")
        self.nob_area = self.area.Cols("no")
        self.pn_area = self.area.Cols("pn")
        self.dp_area = self.area.Cols("dp")
        self.pop_area = self.area.Cols("pop")
        self.name_area = self.area.Cols("name")
        
        # Сечения
        if path_sech != "No" and path_shabl_sech != "No":
            self.rastr.load(0, path_sech, path_shabl_sech)
            self.sech = self.rastr.Tables("sechen")
            self.ns = self.sech.Cols("ns")
            self.psech = self.sech.Cols("psech")
            self.pmax_sech = self.sech.Cols("pmax")
        
        # СХН
        self.polin = self.rastr.Tables("polin")
        self.nsx_polin = self.polin.Cols("nsx")
        
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
    
    def plot_figure(self,name,x,y,char_mode_name):
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
        na_str = input(
            "Введите номера интересующих районов через запятую, либо нажмите 'Enter' для проверки всeх:")
        print("Начинается выполнение, пожалуйста подождите")
        if na_str != "":
            na = [int(x) for x in na_str.split(",")]
        else:
            na = self.give_na(0)
        flag = False
        for n in na:
            print(f"Проверяю район номер {n}")
            # Перебор всех характерных режимов по всем годам
            for mode in self.modes_path:
                # Название характерного режима
                name_char_mode = os.path.basename(mode[0])[:-9]
                k = 0
                for mode_year in mode:
                    # Подгрузка используемого режима
                    self.rastr_work(mode_year)
                    # Выборка и получение значений
                    self.area.SetSel(f"na={n}")
                    ind = self.area.FindNextSel(-1)
                    nm_ar = self.name_area.Z(ind)
                    # Процент потерь от потребления
                    percent = self.dp_area.Z(ind)*100/self.pop_area.Z(ind)
                    # Проверка значения потерь (Dp) 
                    if percent > dp_set:
                        text = (f"""Warning: превышение заданного ({dp_set}%-го) значения потерь от потребления
                                Район: {nm_ar}
                                Характерный режим: {name_char_mode}
                                Год: {self.year_mass[k]}
                                Значение потерь: {round(self.dp_area.Z(ind),2)} МВт
                                Значение потребления: {round(self.pop_area.Z(ind),2)} МВт
                                Потери от потребления: {round(percent,2)}%
                                """)
                        with open(f"{self.path_report}\\report_P_потери.txt", mode="a+") as f:
                                f.write(f"\n{text}")
                        flag = True
                    k += 1
        if flag!=True:
            text = f"Значения потерь в районах корректны"
            with open(f"{self.path_report}\\report_P_потери.txt", mode="a+") as f:
                f.write(f"\n{text}")

            
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 2-го критерия по потреблению

    def crit_2(self, figure):
        na_str = input(
            "Введите номера интересующих районов через запятую, либо нажмите 'Enter' для проверки всeх:")
        print("Начинается выполнение, пожалуйста подождите")
        if na_str != "":
            na = [int(x) for x in na_str.split(",")]
        else:
            na = self.give_na()
        # Получим значения Pпотр для всех районов и всех режимов за все года
        # Перебор всех заданных районов
        flag2 = False
        for n in na:
            print(f"Проверяю район номер {n}")
            # Перебор всех характерных режимов по всем годам
            for mode in self.modes_path:
                Pop_mass_mode = []
                # Название характерного режима
                name_char_mode = os.path.basename(mode[0])[:-9]
                for mode_year in mode:
                    # Подгрузка используемого режима
                    self.rastr_work(mode_year)
                    # Выборка и получение значений
                    self.area.SetSel(f"na={n}")
                    ind = self.area.FindNextSel(-1)
                    Pop_mass_mode.append(self.pop_area.Z(ind))
                nm_ar = self.name_area.Z(ind)
                
                # Проверка в соответсвии с заданными значениями per_up_set и per_down_set
                k = 0   # Предыдущее значение P
                j = 0   # Для счетчика
                for p in Pop_mass_mode:
                    flag = False
                    if k != 0:
                        year_1 = self.year_mass[j-1]
                        year_2 = self.year_mass[j]
                        if round((k-p)*100/k, 2) > per_down_set:
                            text_reason = f"снижение Pпотр на {round((k-p)*100/k, 2)}% ({round(k-p, 1)} МВт)"
                            flag=True
                        if p > (1+per_up_set/100)*k*(year_2-year_1):
                            text_reason = f"увеличение Pпотр на {round((p-k)*100/k, 2)}% ({round(p-k, 1)} МВт)"
                            flag=True
                        if flag:
                            text = (f"""Warning: {text_reason}
                                Район: {nm_ar}
                                Характерный режим: {name_char_mode}
                                Года:
                                        {year_1} - Pпотр = {k}
                                        {year_2} - Pпотр = {p}
                                """)
                            with open(f"{self.path_report}\\report_P_потр.txt", mode="a+") as f:
                                f.write(f"\n{text}")
                            flag2 = True
                    k = p
                    j += 1
                    
                # Строим график если нужно
                if figure:
                    if name_char_mode.find("Зима") != -1:
                        label = name_char_mode[name_char_mode.find("Зима"):]
                    elif name_char_mode.find("Лето") != -1:
                        label = name_char_mode[name_char_mode.find("Лето"):]
                    else:
                        label = name_char_mode[name_char_mode.find("Паводок"):]
                    self.plot_figure(nm_ar, self.year_mass, Pop_mass_mode, label)
            
            plt.savefig(f'{self.path_figure}\\{nm_ar}.png',
                    bbox_inches='tight')
            plt.close()
        if flag2!=True:
            text = f"P_потр монотонно изменяется в заданных пределах"
            with open(f"{self.path_report}\\report_P_потр.txt", mode="a+") as f:
                f.write(f"\n{text}")
            


    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 3-го критерия по расчетному напряжению
    
    def crit_3(self):
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
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
                        u_ras=self.u_ras.Z(index)
                        if u_ras > 0 and (u_ras < (Umin_set/100)*Unom or u_ras > Umax_set[k]):
                            mass_ny.append(self.ny_node.Z(index))
                        index = self.node.FindNextSel(index)
                    if len(mass_ny)>0:
                        text = (f"""Warning: расчетные значения напряжений выходят из заданных пределов
                                    Год: {self.year_mass[j]}
                                    Характерный режим: {name_char_mode}
                                    Uном: {Unom} кВ
                                    Номера узлов: {mass_ny}
                                    """)
                        with open(f"{self.path_report}\\report_U_расч.txt", mode="a+") as f:
                            f.write(f"\n{text}")
                        flag = True
                    k += 1
                j += 1
        if flag!=True:
            text = f"Расчетные значения напряжений не выходят из заданных пределов"
            with open(f"{self.path_report}\\report_U_расч.txt", mode="a+") as f:
                f.write(f"\n{text}")
       
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 4-го критерия по отклонению U_зд от номинального
    
    def crit_4(self):
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
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
                    if u_zd > u_nom*(1+Umax_zd_set/100) or u_zd < u_nom*(1-Umin_zd_set/100):
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.node.FindNextSel(index)
                if len(mass_ny)>0:
                    text = (f"""Warning: заданное напряжение (U_зд) выходит из заданных пределов (-{Umin_zd_set}%; +{Umax_zd_set}%)
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\\report_U_зд.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag!=True:
            text = f"Все заданные напряжения имеют корректные значения"
            with open(f"{self.path_report}\\report_U_зд.txt", mode="a+") as f:
                f.write(f"\n{text}")
                
                
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 5-го критерия по соответсвтвию P и Q нагрузки
    
    def crit_5(self):
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
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
                    tan = qn/pn
                    if tan < tan_min or tan > tan_max:
                        text = (f"""Warning: отношение Q_нагр к P_нагр выходит из заданных пределов
                            Год: {self.year_mass[j]}
                            Характерный режим: {name_char_mode}
                            Номер узла: {self.ny_node.Z(index)}
                            Q_нагр: {qn}
                            P_нагр: {pn}
                            Тангенс: {tan}
                            """)
                        with open(f"{self.path_report}\\report_P_Q_нагр.txt", mode="a+") as f:
                            f.write(f"\n{text}")
                        flag = True
                    index = self.node.FindNextSel(index)
                j += 1
        if flag!=True:
            text = f"Все узлы имеют корректное отношение Q_нагр к P_нагр"
            with open(f"{self.path_report}\\report_P_Q_нагр.txt", mode="a+") as f:
                f.write(f"\n{text}")
    
    
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 6-го критерия по привязке узлов к районам
    
    def crit_6(self):
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
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
                if len(mass_ny)>0:
                    text = (f"""Warning: отсутствует привязка узлов к районам
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\\report_узлы_к_районам.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag!=True:
            text = f"Все узлы привязаны к районам"
            with open(f"{self.path_report}\\report_узлы_к_районам.txt", mode="a+") as f:
                f.write(f"\n{text}")
    
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 7-го критерия заданию номера объединения в районах
    def crit_7(self):
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
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
                if len(mass_na)>0:
                    text = (f"""Warning: отсутствуют номера объединений районов
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера районов: {mass_na}
                                """)
                    with open(f"{self.path_report}\\report_объединения_районов.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag!=True:
            text = f"Все районы имеют заданный номер объединения"
            with open(f"{self.path_report}\\report_объединения_районов.txt", mode="a+") as f:
                f.write(f"\n{text}")
    
    
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 8-го критерия по превышению перетоков в сечениях
    
    def crit_8(self):
        path_sech = input("Задайте путь к файлу сечений:\n")
        path_shabl_sech = input("Задайте путь к шаблону сечений:\n")
        # path_shabl_sech = input("Задайте путь к шаблону сечений:\n")
        flag = False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
            for mode_year in mode:
                self.rastr_work(mode_year, path_sech, path_shabl_sech)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                self.sech.SetSel("pmax>0")
                index = self.sech.FindNextSel(-1)
                # Перебор всех сечений
                mass_ns = []
                for i in range(self.sech.Count):
                    # Проверка условия
                    if self.psech.Z(index) > self.pmax_sech.Z(index):
                        mass_ns.append(self.ns.Z(index))
                    index = self.sech.FindNextSel(index)
                if len(mass_ns)>0:
                    text = (f"""Warning: превышение перетоков в сечениях
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера сечений: {mass_ns} 
                                """)
                    with open(f"{self.path_report}\\report_перетоки_сечений.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag = True
                j += 1
        if flag!=True:
            text = f"Все сечения имеют корректные перетоки"
            with open(f"{self.path_report}\\report_перетоки_сечений.txt", mode="a+") as f:
                f.write(f"\n{text}")
    
    
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 9-го критерия по неизменности количества строк в таблицах
    def crit_9(self):
        for i in range(len(self.year_mass)):
            print(f"Проверяю {self.year_mass[i]} год")
            with open(f"{self.path_report}\\report_количество_строк.txt", mode="a+") as f:
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
                with open(f"{self.path_report}\\report_количество_строк.txt", mode="a+") as f:
                    f.write(f"\n{amount_mass} - {os.path.basename(self.modes_path[k][i])}")
    
    
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 10-го критерия по соответствию используемых СХН
    def crit_10(self):
        flag2=False
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:-4]
                # Номера имеющихся СХН
                mass_nsx = []
                for k in range(self.polin.Count):
                    mass_nsx.append(self.nsx_polin.Z(k))
                # Перебор всех узлов
                mass_ny = []
                self.node.SetSel("nsx>=1")
                index = self.node.FindNextSel(-1)
                for i in range(self.node.Count):
                    flag=False
                    # Проверка условия
                    for nsx in mass_nsx:
                        # Если СХН узла есть в таблице СХН, то узел добавляется в массив
                        if self.nsx_node.Z(index)==nsx:
                            flag=True
                    if flag:
                        mass_ny.append(self.ny_node.Z(index))
                    index = self.area.FindNextSel(index)
                if len(mass_ny)>0:
                    text = (f"""Warning: обращение к несуществующему СХН
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номера узлов: {mass_ny}
                                """)
                    with open(f"{self.path_report}\\report_СХН_узлы.txt", mode="a+") as f:
                        f.write(f"\n{text}")
                    flag2 = True
                j += 1
        if flag2!=True:
            text = f"СХН во всех узлах заданны верно"
            with open(f"{self.path_report}\\report_СХН_узлы.txt", mode="a+") as f:
                f.write(f"\n{text}")

def start(direct_path):
    direct_path = input("Введите путь к папкам годов:\n")
    print("Очистить данные отчетов?")
    cl_rep = input("Напишите 'Yes' если да; нажмите 'Enter' если нет: ")
    clear_rep = cl_rep == 'Yes'
    print("Очистить папку с графиками?")
    cl_fig = input("Напишите 'Yes' если да; нажмите 'Enter' если нет: ")
    clear_figure = cl_fig == 'Yes'
    print("Строить графики?")
    fig = input("Нажмите 'Enter' если да; напишите 'No' если нет: ")
    figure = fig ==''
    print("""Номера критериев:
    1 - Проверка корректности заданных потерь Района;
    2 - Проверка немонотонности изменения потребления Района;
    3 - Проверка отклонения расчетного напряжения;
    4 - Проверка отклонения U_зд от номинального;
    5 - Проверка соотношения P и Q нагрузки;
    6 - Проверка привязки узлов к районам;
    7 - Проверка задания номера объединения в таблице "Районы";
    8 - Проверка превышения перетоков в сечениях;
    9 - Проверка неизменности количества строк в характерных режимах;
    10 - Проверка взаимосвязи заданных СХН в узлах с таблицей СХН (НЕ ГОТОВО)
          """)
    num_crit = input("Введите номер критерия: ")
    match num_crit:
        case "1":
            CheckUp(direct_path, clear_rep, clear_figure).crit_1()
        case "2":
            CheckUp(direct_path, clear_rep, clear_figure).crit_2(figure)
        case "3":
            CheckUp(direct_path, clear_rep, clear_figure).crit_3()
        case "4":
            CheckUp(direct_path, clear_rep, clear_figure).crit_4()
        case "5":
            CheckUp(direct_path, clear_rep, clear_figure).crit_5()
        case "6":
            CheckUp(direct_path, clear_rep, clear_figure).crit_6()
        case "7":
            CheckUp(direct_path, clear_rep, clear_figure).crit_7()
        case "8":
            CheckUp(direct_path, clear_rep, clear_figure).crit_8()
        case "9":
            CheckUp(direct_path, clear_rep, clear_figure).crit_9()
        # case "10":
        #     CheckUp(direct_path, clear_rep, clear_figure).crit_10()
        case _:
            print("Неверный критерий")
    return direct_path

path = ''
while True:
    flag = input("Чтобы начать новый расчет нажмите 'Enter'")
    if flag == "":
        path = start(path)
    else:
        break

# dir_path = r'C:\Users\bukre\OneDrive\Рабочий стол\ОДУ\Проверка расчетных схем\02_ПРМ с учетом корректировок'
# clear_rep = True
# clear_figure = clear_rep
# a = CheckUp(dir_path, clear_rep, clear_figure).crit_2(figure=True)
# print(a)

# C:\Users\bukre\OneDrive\Рабочий стол\ОДУ\Проверка расчетных схем\Файл сечений\сечения.sch