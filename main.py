import os
import pandas as pd
import tkinter.messagebox
import numpy as np
import xlwt
import datetime

now_time = datetime.datetime.now()


def getFiles(sourceDir):
    listfile = []
    for file in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir, file)
        if os.path.isfile(sourceFile):
            listfile.append(sourceFile)
    return listfile


class filter_T:

    def __init__(self, data):
        self.data = data
        self.min_t_data = self.data['最低温度']
        self.size = len(self.min_t_data) - 2
        self.course = []
        self.year = set()

        for i in range(len(self.min_t_data)):
            self.get_year(self.data['日期'][i])
            if len(self.course) > 0:
                if self.course[-1]['data'][-1] > i:
                    continue

            if self.down_24hour(i):
                if min(self.min_t_data[i:i + 2]) < 15:
                    self.course.append({'f': '24', 'data': (i, i + 2), 'days': 2})
                # for j in range(i + 1, self.size):
                #     if self.min_t_data[j] < self.min_t_data[j - 1]:
                #         continue
                #     else:
                #         if j - i > 1:
                #             if min(self.min_t_data[i:j]) < 15:
                #                 self.course.append({'f': '24', 'data': (i, j), 'days': j - i})
                #         break
            elif self.down_48hour(i):
                if min(self.min_t_data[i:i + 3]) < 15:
                    self.course.append({'f': '48', 'data': (i, i + 3), 'days': 3})
                # for j in range(i + 1, self.size):
                #     if self.min_t_data[j] < self.min_t_data[j - 1]:
                #         continue
                #     else:
                #         if j - i > 2:
                #             if min(self.min_t_data[i:j]) < 15:
                #                 self.course.append({'f': '48', 'data': (i, j), 'days': j - i})
                #         break

            elif self.down_72hour(i):
                if min(self.min_t_data[i:i + 4]) < 15:
                    self.course.append({'f': '72', 'data': (i, i + 4), 'days': 4})
                # for j in range(i + 1, self.size):
                #     if self.min_t_data[j] < self.min_t_data[j - 1]:
                #         continue
                #     else:
                #         if j - i > 2:
                #             if min(self.min_t_data[i:j]) < 15:
                #                 self.course.append({'f': '72', 'data': (i, j), 'days': j - i})
                #         break

    def down_24hour(self, num):

        if self.size <= num + 1:
            return False
        if self.min_t_data[num] == np.inf or self.min_t_data[num + 3] == np.inf:
            return False
        if self.min_t_data[num] < self.min_t_data[num + 1]:
            return False
        return True if self.min_t_data[num] - self.min_t_data[num + 1] > 8 else False

    def down_48hour(self, num):

        if self.size <= num + 2:
            return False
        if self.min_t_data[num] == np.inf or self.min_t_data[num + 3] == np.inf:
            return False
        if self.min_t_data[num] < self.min_t_data[num + 1] or self.min_t_data[num + 1] < self.min_t_data[num + 2]:
            return False
        return True if self.min_t_data[num] - self.min_t_data[num + 2] > 10 else False

    def down_72hour(self, num):
        if self.size <= num + 3:
            return False
        if self.min_t_data[num] == np.inf or self.min_t_data[num + 3] == np.inf:
            return False
        if self.min_t_data[num] < self.min_t_data[num + 1] or self.min_t_data[num + 1] < self.min_t_data[num + 2] or \
                self.min_t_data[num + 2] < self.min_t_data[num + 3]:
            return False
        return True if self.min_t_data[num] - self.min_t_data[num + 3] > 12 else False

    def get_year(self, date):
        try:
            self.year.add(date[:4])
        except:
            pass


out_data = []
for file in getFiles('.\\data'):
    ed = pd.read_excel(file)
    ed = ed.fillna(np.inf)
    ed_li = ed.values[:, 3]
    ed_li = ed_li.astype(np.float32)
    ed['最低温度'] = ed_li

    ed.loc[ed.shape[0]] = [None, np.inf, np.inf, np.inf, np.inf]

    c = filter_T(ed)
    ed = ed.values
    line = 1
    # 创建工作簿对象
    work_book = xlwt.Workbook()
    # 创建工作表对象
    st = work_book.add_sheet("sheet1")
    st.write(0, 0, '日期')
    st.write(0, 1, '平均温度')
    st.write(0, 2, '最高温度')
    st.write(0, 3, '最低温度')
    st.write(0, 4, '日温差')
    st.write(0, 6, '过程')
    for i in range(len(c.course)):
        for j in range(c.course[i]['data'][0], c.course[i]['data'][1]):
            st.write(line + 1, 0, ed[j, 0])
            st.write(line + 1, 1, ed[j, 1])
            st.write(line + 1, 2, ed[j, 2])
            st.write(line + 1, 3, ed[j, 3])
            st.write(line + 1, 4, ed[j, 4])
            st.write(line + 1, 6, '过程' + str(i + 1))
            line += 1
        line += 1

    work_book.save('.\\out\\' + file[7:-4] + '(out).xls')

    # 制作年份分类字典
    year_dict = {}
    for i in c.year:
        year_dict[i] = {'all': []}
    for i in c.course:
        for j in c.year:
            if str(j) in ed[:, 0][i['data'][0]]:
                year_dict[j]['all'].append(i)

    for i in c.year:
        try:
            year_dict[i]['mean_day'] = sum(j['days'] for j in year_dict[i]['all']) / len(year_dict[i]['all'])
        except:
            year_dict[i]['mean_day'] = 0
        year_dict[i]['file'] = file[7:-4]
        min_li = []
        meanLi = []
        max_li = []
        days = []
        for k in year_dict[i]['all']:
            max_var = max([ed_li[p] for p in range(k['data'][0], k['data'][-1])])
            min_var = min([ed_li[p] for p in range(k['data'][0], k['data'][-1])])
            days.append(k['days'])
            min_li.append(min_var)
            max_li.append(max_var)
            meanLi.append(max_var - min_var)
        year_dict[i]['dayNum'] = ','.join('%s' % i for i in days)
        year_dict[i]['gcNum'] = len(year_dict[i]['all'])
        try:
            year_dict[i]['max_min_var'] = sum(meanLi) / len(year_dict[i]['all'])
        except:
            year_dict[i]['max_min_var'] = 'None'
        try:
            year_dict[i]['min_mean'] = sum(min_li) / len(year_dict[i]['all'])
        except:
            year_dict[i]['min_mean'] = 'None'
        try:
            year_dict[i]['max_mean'] = sum(max_li) / len(year_dict[i]['all'])
        except:
            year_dict[i]['max_mean'] = 'None'
    year_dict['year_data'] = c.year
    # print(year_dict, '\n\n')

    out_data.append(year_dict)

out_xls = xlwt.Workbook()
st = out_xls.add_sheet("sheet1")
st.write(0, 0, '站点')
st.write(0, 1, '年份')
st.write(0, 2, '过程平均天数')
st.write(0, 3, 'MAX-MIN平均')
st.write(0, 4, 'MIN平均')
st.write(0, 5, 'MAX平均')
st.write(0, 6, '天数')
st.write(0, 7, '过程数')

line = 0
for i in range(len(out_data)):
    for y in sorted([str(h) for h in out_data[i]['year_data']]):
        if out_data[i][y]['max_min_var'] == 'None':
            continue
        # 这里是写入数据的地方
        st.write(line + 1, 0, out_data[i][y]['file'])
        st.write(line + 1, 1, y)
        st.write(line + 1, 2, out_data[i][y]['mean_day'])
        st.write(line + 1, 3, out_data[i][y]['max_min_var'])
        st.write(line + 1, 4, out_data[i][y]['min_mean'])
        st.write(line + 1, 5, out_data[i][y]['max_mean'])
        # 下面是过程数和天数，不需要可以删掉
        st.write(line + 1, 6, out_data[i][y]['dayNum'])
        st.write(line + 1, 7, out_data[i][y]['gcNum'])
        line += 1

try:
    out_xls.save('.\\statistics_out\\out.xls')
except PermissionError as e:
    tkinter.messagebox.showinfo('错误!', 'xls文件未关闭！')

l_time = datetime.datetime.now()

print('\n耗时：' + str(l_time - now_time))
