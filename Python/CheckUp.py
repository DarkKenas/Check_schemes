import os
import win32com.client
import matplotlib.pyplot as plt
from config import pn_set, per_up_set, per_down_set, dp_set, Unom_set, Umax_set, Umin_set


class CheckUp:
    def __init__(self, directory_path: str, clear_report: bool, clear_figure: bool):
        # ------------------------------------------------------------------
        # Создаём папку отчетных файлов программы если её нет
        self.path_report = f"./Отчетные файлы"
        self.path_figure = f"./Отчетные файлы/Графики"
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

    def rastr_work(self, path_mode):
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
        self.nsx = self.node.Cols("nsx")
        self.pn_node = self.node.Cols("pn")

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

        # Территории
        # Объединения

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод получения всех районов, где Pнаг > pn_set (заданного)

    def give_na(self, pn=pn_set):
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
                    k += 1

            
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
                    flag=False
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

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Метод для проверки 3-го критерия по напряжению
    
    def crit_3(self):
        for mode in self.modes_path:
            print("Проверяю характерный режим")
            j=0
            for mode_year in mode:
                self.rastr_work(mode_year)
                # Название характерного режима
                name_char_mode = os.path.basename(mode_year)[:4]
                k = 0
                # Перебор всех Uном
                for Unom in Unom_set:
                    self.node.SetSel(f"uhom={Unom}")
                    index = self.node.FindNextSel(-1)
                    # Перебор всех узлов с рассматриваемым Uном
                    for i in range(self.node.Count):
                        # Проверка условия
                        u_ras=self.u_ras.Z(index)
                        if u_ras > 0 and (u_ras < (Umin_set/100)*Unom or u_ras > Umax_set[k]):
                            text = (f"""Warning: расчетное значение напряжения выходит из заданных пределов
                                Год: {self.year_mass[j]}
                                Характерный режим: {name_char_mode}
                                Номер узла: {self.ny_node.Z(index)}
                                Uном: {Unom}
                                Uрасч: {self.u_ras.Z(index)}
                                """)
                            with open(f"{self.path_report}\\report_U_расч.txt", mode="a+") as f:
                                    f.write(f"\n{text}")
                        index = self.node.FindNextSel(index)
                    k += 1
                j += 1
                            
                        
    
    
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
    1 - Проверка корректности заданных потерь Района
    2 - Проверка немонотонности изменения потребления Района
    3 - Проверка отклонения расчетного напряжения
          """)
    num_crit = input("Введите номер критерия: ")
    match num_crit:
        case "1":
            CheckUp(direct_path, clear_rep, clear_figure).crit_1()
        case "2":
            CheckUp(direct_path, clear_rep, clear_figure).crit_2(figure)
        case "3":
            CheckUp(direct_path, clear_rep, clear_figure).crit_3()
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
