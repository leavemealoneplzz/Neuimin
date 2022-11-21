import csv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side
import matplotlib.pyplot as plt
import numpy as np


class Vacancies:
    currency = {
        "AZN": 35.68, "BYR": 23.91, "EUR": 59.90, "GEL": 21.74, "KGS": 0.76,
        "KZT": 0.13, "RUR": 1, "UAH": 1.64, "USD": 60.66, "UZS": 0.0055,
    }
    def __init__(self, vacancy):
        self.name = vacancy['name']
        self.salary_from = int(float(vacancy['salary_from']))
        self.salary_to = int(float(vacancy['salary_to']))
        self.salary_currency = vacancy['salary_currency']
        self.salary_average = self.currency[self.salary_currency] * (self.salary_from + self.salary_to) / 2
        self.area_name = vacancy['area_name']
        self.year = int(vacancy['published_at'][:4])


class Data:
    def __init__(self, file_name, vacancy_name):
        self.file_name = file_name
        self.vacancy_name = vacancy_name

    @staticmethod
    def incrementation(dic, key, amount):
        if key in dic:
            dic[key] += amount
        else:
            dic[key] = amount

    @staticmethod
    def avg(dic):
        new_dic = {}
        for key, values in dic.items():
            new_dic[key] = int(sum(values) / len(values))
        return new_dic

    @staticmethod
    def stats_print(stats1, stats2, stats3, stats4, stats5, stats6):
        print('Динамика уровня зарплат по годам: {0}'.format(stats1))
        print('Динамика количества вакансий по годам: {0}'.format(stats2))
        print('Динамика уровня зарплат по годам для выбранной профессии: {0}'.format(stats3))
        print('Динамика количества вакансий по годам для выбранной профессии: {0}'.format(stats4))
        print('Уровень зарплат по городам (в порядке убывания): {0}'.format(stats5))
        print('Доля вакансий по городам (в порядке убывания): {0}'.format(stats6))

    def vac_increments(self, count_vac, salary, salary_city, salary_vac_name):
        for vac_dic in self.csv_r():
            vacancy = Vacancies(vac_dic)
            self.incrementation(salary, vacancy.year, [vacancy.salary_average])
            if vacancy.name.find(self.vacancy_name) != -1:
                self.incrementation(salary_vac_name, vacancy.year, [vacancy.salary_average])
            self.incrementation(salary_city, vacancy.area_name, [vacancy.salary_average])
            count_vac += 1
        return count_vac

    def get_stats(self, count_vac, salary, salary_city, salary_vac_name):
        stats = self.avg(salary)
        stats2 = self.avg(salary_vac_name)
        stats3 = self.avg(salary_city)
        stats4 = {}
        for year, salaries in salary_city.items():
            stats4[year] = round(len(salaries) / count_vac, 4)
        stats4 = list(filter(lambda a: a[-1] >= 0.01, [(key, value) for key, value in stats4.items()]))
        stats4.sort(key=lambda a: a[-1], reverse=True)
        stats5 = stats4.copy()
        stats4 = dict(stats4)
        stats3 = list(filter(lambda a: a[0] in list(stats4.keys()), [(key, value) for key, value in stats3.items()]))
        stats3.sort(key=lambda a: a[-1], reverse=True)
        stats3 = dict(stats3[:10])
        stats5 = dict(stats5[:10])
        return stats, stats2, stats3, stats5

    def csv_r(self):
        with open(self.file_name, mode='r', encoding='utf-8-sig') as file:
            read = csv.reader(file)
            head = next(read)
            header_len = len(head)
            for row in read:
                if '' not in row and len(row) == header_len:
                    yield dict(zip(head, row))

    def res_stats(self):
        salary = {}
        salary_vac_name = {}
        salary_city = {}
        count_vac = 0

        count_vac = self.vac_increments(count_vac, salary, salary_city, salary_vac_name)

        vac_num = dict([(key, len(value)) for key, value in salary.items()])
        vac_num_name = dict([(key, len(value)) for key, value in salary_vac_name.items()])

        if not salary_vac_name:
            salary_vac_name = dict([(key, [0]) for key, value in salary.items()])
            vac_num_name = dict([(key, 0) for key, value in vac_num.items()])

        stats, stats2, stats3, stats5 = self.get_stats(count_vac, salary, salary_city, salary_vac_name)

        return stats, vac_num, stats2, vac_num_name, stats3, stats5



class Input:
    def __init__(self):
        self.file_name = input('Введите название файла: ')
        self.vacancy_name = input('Введите название профессии: ')

        dataset = Data(self.file_name, self.vacancy_name)
        stats1, stats2, stats3, stats4, stats5, stats6 = dataset.res_stats()
        dataset.stats_print(stats1, stats2, stats3, stats4, stats5, stats6)

        report = Report(self.vacancy_name, stats1, stats2, stats3, stats4, stats5, stats6)
        report.exc_gen()
        report.saving('report.xlsx')
        report.img_gen()


class Report:
    def __init__(self, vacancy_name, stats1, stats2, stats3, stats4, stats5, stats6):
        self.wb = Workbook()
        self.vacancy_name = vacancy_name
        self.stats1 = stats1
        self.stats2 = stats2
        self.stats3 = stats3
        self.stats4 = stats4
        self.stats5 = stats5
        self.stats6 = stats6

    def exc_gen(self):
        ws = self.wb.active
        ws.title = 'Статистика по годам'
        ws.append(['Год', 'Средняя зарплата', 'Средняя зарплата - ' + self.vacancy_name, 'Количество вакансий', 'Количество вакансий - ' + self.vacancy_name])
        for year in self.stats1.keys():
            ws.append([year, self.stats1[year], self.stats3[year], self.stats2[year], self.stats4[year]])

        data = [['Год ', 'Средняя зарплата ', ' Средняя зарплата - ' + self.vacancy_name, ' Количество вакансий', ' Количество вакансий - ' + self.vacancy_name]]
        col_width = []
        for row in data:
            for i, cell in enumerate(row):
                if len(col_width) > i:
                    if len(cell) > col_width[i]:
                        col_width[i] = len(cell)
                else:
                    col_width += [len(cell)]

        for i, col_width in enumerate(col_width, 1):
            ws.column_dimensions[get_column_letter(i)].width = col_width + 2

        data = []
        data.append(['Город', 'Уровень зарплат', '', 'Город', 'Доля вакансий'])
        for (city1, value1), (city2, value2) in zip(self.stats5.items(), self.stats6.items()):
            data.append([city1, value1, '', city2, value2])
        ws2 = self.wb.create_sheet('Статистика по городам')
        for row in data:
            ws2.append(row)

        col_width = self.col_width_res(col_width, data)

        for i, col_width in enumerate(col_width, 1):  # ,1 to start at 1
            ws2.column_dimensions[get_column_letter(i)].width = col_width + 2

        fontb = Font(bold=True)
        for col in 'ABCDE':
            ws[col + '1'].font = fontb
            ws2[col + '1'].font = fontb

        for index, _ in enumerate(self.stats5):
            ws2['E' + str(index + 2)].number_format = '0.00%'

        thin = Side(border_style='thin', color='00000000')

        for row in range(len(data)):
            for col in 'ABDE':
                ws2[col + str(row + 1)].border = Border(left=thin, bottom=thin, right=thin, top=thin)

        for row, _ in enumerate(self.stats1):
            for col in 'ABCDE':
                ws[col + str(row + 1)].border = Border(left=thin, bottom=thin, right=thin, top=thin)


    def col_width_res(self, col_width, data):
        col_width = []
        for row in data:
            for i, cell in enumerate(row):
                cell = str(cell)
                if len(col_width) > i:
                    if len(cell) > col_width[i]:
                        col_width[i] = len(cell)
                else:
                    col_width += [len(cell)]
        return col_width


    def img_gen(self):
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(nrows=2, ncols=2)

        bar1 = ax1.bar(np.array(list(self.stats1.keys())) - 0.4, self.stats1.values(), width=0.4)
        bar2 = ax1.bar(np.array(list(self.stats1.keys())), self.stats3.values(), width=0.4)
        ax1.set_title('Уровень зарплат по годам', fontdict={'fontsize': 8})
        ax1.grid(axis='y')
        ax1.legend((bar1[0], bar2[0]), ('средняя з/п', 'з/п ' + self.vacancy_name.lower()), prop={'size': 8})
        ax1.set_xticks(np.array(list(self.stats1.keys())) - 0.2, list(self.stats1.keys()), rotation=90)
        ax1.xaxis.set_tick_params(labelsize=8)
        ax1.yaxis.set_tick_params(labelsize=8)

        ax2.set_title('Количество вакансий по годам', fontdict={'fontsize': 8})
        bar1 = ax2.bar(np.array(list(self.stats2.keys())) - 0.4, self.stats2.values(), width=0.4)
        bar2 = ax2.bar(np.array(list(self.stats2.keys())), self.stats4.values(), width=0.4)
        ax2.legend((bar1[0], bar2[0]), ('Количество вакансий', 'Количество вакансий\n' + self.vacancy_name.lower()), prop={'size': 8})
        ax2.set_xticks(np.array(list(self.stats2.keys())) - 0.2, list(self.stats2.keys()), rotation=90)
        ax2.grid(axis='y')
        ax2.xaxis.set_tick_params(labelsize=8)
        ax2.yaxis.set_tick_params(labelsize=8)

        ax3.set_title('Уровень зарплат по городам', fontdict={'fontsize': 8})
        ax3.barh(list([str(a).replace(' ', '\n').replace('-', '-\n') for a in reversed(list(self.stats5.keys()))]), list(reversed(list(self.stats5.values()))), color='blue', height=0.5, align='center')
        ax3.yaxis.set_tick_params(labelsize=6)
        ax3.xaxis.set_tick_params(labelsize=8)
        ax3.grid(axis='x')

        ax4.set_title('Доля вакансий по городам', fontdict={'fontsize': 8})
        other = 1 - sum([value for value in self.stats6.values()])
        ax4.pie(list(self.stats6.values()) + [other], labels=list(self.stats6.keys()) + ['Другие'], textprops={'fontsize': 6})

        plt.tight_layout()
        plt.savefig('graph.png')

    def saving(self, filename):
        self.wb.save(filename=filename)


if __name__ == '__main__':
    Input()