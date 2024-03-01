import csv
import pandas as pd
import numpy as np
import datetime
import random

#####################################################################################
# 使用者上傳excel
from tkinter import *
from tkinter import filedialog

def open_file():  # 讓使用者上傳excel檔
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if filename:
        process_file(filename)

def process_file(filename):  # 處理excel檔案
    data = pd.read_excel(filename, sheet_name=None)
    for sheet_name, df in data.items():
        # 依序讀入4張sheet
        if sheet_name == "project":
            project = df
            project = project.set_axis(project.iloc[:, 0], axis=0, copy=False) # 換掉列名稱
            project = project.iloc[:, 1:]  # 移除第一欄
            #print(project)
        if sheet_name == "level":
            level = df
            level = level.set_axis(level.iloc[:, 0], axis=0, copy=False) # 換掉列名稱
            level = level.iloc[:, 1:]  # 移除第一欄
            #print(level)
        if sheet_name == "level_and_module":
            level_and_module = df
            level_and_module = level_and_module.set_axis(level_and_module.iloc[:, 0], axis=0, copy=False) # 換掉列名稱
            level_and_module = level_and_module.iloc[:, 1:]  # 移除第一欄
            #print(level_and_module)
        elif sheet_name == "project_and_module":
            project_and_module = df
            project_and_module = project_and_module.set_axis(project_and_module.iloc[:, 0], axis=0, copy=False) # 換掉列名稱
            project_and_module = project_and_module.iloc[:, 1:]  # 移除第一欄
            #print(project_and_module)
    
    ###################################################
    # 讓使用者輸入PM最多做多少
    param_label = Label(window, text="PM最多做多少人天(請填數字):")
    param_label.pack()

    param_entry = Entry(window)
    param_entry.pack()
    
    # 創建確認按鈕，點擊時儲存參數值並禁用輸入框
    confirm_button = Button(window, text="確定", command=lambda: confirm_param(param_entry, project, level, level_and_module, project_and_module))
    confirm_button.pack()
    
    ######################################################
    
       
def confirm_param(param_entry, project, level, level_and_module, project_and_module):
    global param_value
    param_value = param_entry.get()
    param_entry.config(state=DISABLED)

    # 進入主要演算法 並將excel檔4個sheet匯入
    PMs_best_schedule_df, project_level_days, project_level_days_cost, PM_schedule, L4_schedule, L3_schedule, L2_schedule, L1_schedule = main_function(project, level, level_and_module, project_and_module, param_value)
    if type(PMs_best_schedule_df) == int:
        print('此段期間專案做不完')
        return '此段期間專案做不完'
    
    # 創建下載按鈕
    # 創建按鈕 1（下載 Excel1: for financial）
    download_button1 = Button(window, text="Download financial_data", command=lambda:download_excel_financial_data(PMs_best_schedule_df, project_level_days, project_level_days_cost))
    download_button1.pack()

    # 創建按鈕 2（下載 Excel 2: for PM scheduling）
    download_button2 = Button(window, text="Download PM_scheduling", command=lambda:download_excel_scheduling_data(PM_schedule))
    download_button2.pack()

    # 創建按鈕 3（下載 Excel 3: for L4 scheduling）
    download_button3 = Button(window, text="Download L4_scheduling", command=lambda:download_excel_scheduling_data(L4_schedule))
    download_button3.pack()

    # 創建按鈕 4（下載 Excel 4: for L3 scheduling）
    download_button4 = Button(window, text="Download L3_scheduling", command=lambda:download_excel_scheduling_data(L3_schedule))
    download_button4.pack()

    # 創建按鈕 5（下載 Excel 5: for L2 scheduling）
    download_button5 = Button(window, text="Download L2_scheduling", command=lambda:download_excel_scheduling_data(L2_schedule))
    download_button5.pack()

    # 創建按鈕 6（下載 Excel 6: for L1 scheduling）
    download_button6 = Button(window, text="Download L1_scheduling", command=lambda:download_excel_scheduling_data(L1_schedule))
    download_button6.pack()

    # 關閉主視窗
    #window.destroy()

#######################################################################################
# 進入正式演算法

def main_function(project, level, level_and_module, project_and_module, param_PM):
    param_PM = int(param_PM)
    # 整理各專案開始月/結束月(不會動--Global)
    project_time_dict = {}  # 專案-起始/結束
    earlist_month = 12  # 7
    latest_month = 1    # 12
    for name in project.columns:
        project_time_dict[name] = [(project.loc['start_date', name]).month, (project.loc['end_date', name]).month]
    for name in project_time_dict:
        if project_time_dict[name][0] < earlist_month:
            earlist_month =  project_time_dict[name][0]
        elif project_time_dict[name][1] > latest_month:
            latest_month =  project_time_dict[name][1]
    # print(project_time_dict)
    # print(earlist_month, latest_month)

    # 以月當key
    month_project = {}  # 月-專案
    for i in range(earlist_month, latest_month + 1): # 7-12
        for name in project.columns:  # 專案一~六
            if i >= project_time_dict[name][0] and i <= project_time_dict[name][1]:  # 表示該專案進行中
                if i in month_project:
                    month_project[i].append(name)
                else:
                    month_project[i] = [name]
    # print(month_project)

    #####################################################################
    # 假設起點都是1號，結束都是30號
    # 先排加到滿30天
    # 分PM
    import itertools
    PMs = list(level_and_module.index)  # PM可挑的各專長



    # 先決定要抽幾round
    choose_round = {}  # 第幾round有哪幾個專案 每一新的round就會重新抽

    # 結束月份
    endMonths = []  # 9, 12, 11
    for key in project_time_dict:
        if project_time_dict[key][1] not in endMonths:
            endMonths.append(project_time_dict[key][1])
    endMonths.sort()  # 9, 11, 12 ->有可能會抽3輪

    # 決定幾round
    for e in range(len(endMonths)):  # 9月前開始的專案/9-11開始的專案/11-12開始的專案
        for key in project_time_dict:
            if e == 0:
                if project_time_dict[key][0] <= endMonths[e]:
                    if endMonths[e] not in choose_round:
                        choose_round[endMonths[e]] = [key]
                    else:
                        choose_round[endMonths[e]].append(key)
            else:
                if project_time_dict[key][0] <= endMonths[e] and project_time_dict[key][0] > endMonths[e - 1]:
                    if endMonths[e] not in choose_round:
                        choose_round[endMonths[e]] = [key]
                    else:
                        choose_round[endMonths[e]].append(key)
    rounds = len(choose_round)  # 2 rounds
    # print(choose_round)
        
    for r in range(rounds):  # 0 1
        if r == 0:  # 第1round
            num_selections = len(choose_round[endMonths[r]])
            permutations = list(itertools.permutations(PMs, num_selections))  # 取5個排列
            #print(permutations)  # 一二三四六(120種)
        else:  # 第2round (2種)
            # 更新到最新可用的PM
            permutations1 = []  # 120種+120種
            for p1 in permutations:
                PMs = ['PP', 'MM', 'SD', 'CO', 'FI']
                PMs.remove(p1[1]) 
                PMs.remove(p1[3])
                PMs.remove(p1[4]) 
                combinations = list(itertools.combinations(PMs, 1))  # 取出1個
                
                copied_p1 = tuple(p1)  # 複製一個
                copied_p1 += combinations[0]  # 加上去
                permutations1.append(copied_p1)
                
                copied_p12 = tuple(p1)  # 複製一個
                copied_p12 += combinations[1]
                permutations1.append(copied_p12)
    # print(permutations1)

    # 以上僅處理excel資料(excel資料不變，上面資料就不變)
    ##################################################################################
    # 正式進入240種迴圈

    best_solution = False  # 還未找到最佳解
    def find_best_schedule(project, level, level_and_module, project_and_module, project_time_dict, earlist_month, latest_month, permutations1, choose_round, month_project, best_solution):
        feasibility_num = 0  # 有幾種可行解(做得完)
        most_saving_schedule = []  # PM組合--list包tuple
        

        for p in range(len(permutations1)):  # p有240種
            
            #### 每種PM情況都會重新清空上次情況#######################################
            # 剩餘的人力level_and_module
            level_and_module_remain = level_and_module.copy()  # 複製DataFrame
            # print(level_and_module_remain)


            # finish for PM
            project_and_module_finish_PM = project_and_module.copy()  # 複製DataFrame
            project_and_module_finish_PM = pd.DataFrame(np.zeros_like(project_and_module_finish_PM), index=project_and_module_finish_PM.index, columns=project_and_module_finish_PM.columns)
            # print(project_and_module_finish_L4)

            # finish for L4
            project_and_module_finish_L4 = project_and_module.copy()  # 複製DataFrame
            project_and_module_finish_L4 = pd.DataFrame(np.zeros_like(project_and_module_finish_L4), index=project_and_module_finish_L4.index, columns=project_and_module_finish_L4.columns)
            # print(project_and_module_finish_L4)

            # finish for L3
            project_and_module_finish_L3 = project_and_module.copy()  # 複製DataFrame
            project_and_module_finish_L3 = pd.DataFrame(np.zeros_like(project_and_module_finish_L3), index=project_and_module_finish_L3.index, columns=project_and_module_finish_L3.columns)
            # print(project_and_module_finish_L3)

            # finish for L2
            project_and_module_finish_L2 = project_and_module.copy()  # 複製DataFrame
            project_and_module_finish_L2 = pd.DataFrame(np.zeros_like(project_and_module_finish_L2), index=project_and_module_finish_L2.index, columns=project_and_module_finish_L2.columns)
            # print(project_and_module_finish_L2)

            # finish for L1
            project_and_module_finish_L1 = project_and_module.copy()  # 複製DataFrame
            project_and_module_finish_L1 = pd.DataFrame(np.zeros_like(project_and_module_finish_L1), index=project_and_module_finish_L1.index, columns=project_and_module_finish_L1.columns)
            # print(project_and_module_finish_L1)
            
            # dict for finishing key:月份 value:該月finish累計
            project_and_module_finish_PM_dict = {}
            project_and_module_finish_L4_dict = {}
            project_and_module_finish_L3_dict = {}
            project_and_module_finish_L2_dict = {}
            project_and_module_finish_L1_dict = {}

            # 剩餘人天project_and_module
            project_and_module_remain = project_and_module.copy()  # 複製DataFrame
            ####################################################################
         
            
            # 將240種可能性一個一個加入字典
            PMs_each_project = {}  # 專案：PM專長


            for r in range(len(list(choose_round.values()))):
                for e in range(len(list(choose_round.values())[r])):
                    name = list(choose_round.values())[r][e]
                    PMs_each_project[name] = permutations1[p][e]

            # print(PMs_each_project)

            ########################################
            # 分PM的人天
            

            # 計算各level在各專案做了多少人天(用來算成本)
            project_level_days = pd.DataFrame(0, index=project_and_module_remain.columns, columns=level.columns)
            project_level_days_norm = pd.DataFrame(0, index=project_and_module_remain.columns, columns=level.columns)  # 正常
            project_level_days_add = pd.DataFrame(0, index=project_and_module_remain.columns, columns=level.columns)  # 加班
            # print(project_level_days)

            # 計算各level的扣打
            level_idle = pd.DataFrame(columns=['Level', 'Month', 'Module', 'Idle'])  # 空的
            # print(level_idle)

            ###############################第一步#################################################
            # 分PM人天(正式)：一定要做完
            for i in range(earlist_month, latest_month + 1): # 7-12
                # 更新到最新可用的人力->每月都要先更新(專案結束要把人力丟出來)
                for key in project_time_dict:  # 專案一~六
                    if project_time_dict[key][1] + 1 == i:  # 上個月結束，人力(不能跨專案)交出來
                        level_and_module_remain.loc[PMs_each_project[key], 'L4'] += 1
                
                # 分PM
                for name in month_project[i]:  # 該月進行的專案
                    prof = PMs_each_project[name]  # 該專案的PM專長
                    L4_lower_limit = project.loc["project_days", name] * 0.25  # 該專案L4下限
                    
                    # 只有第一個月要扣掉人力(不能跨專案)
                    if project_time_dict[name][0] == i:  # 專案的第一個月
                        level_and_module_remain.loc[prof, 'L4'] -= 1  # 用掉人力(PM)
                    
                    # 先做PM 有餘力再做專長
                    if project_and_module_remain.loc['PM', name] < param_PM:  # 可以做專長
                        finish_pm = project_and_module_remain.loc['PM', name] # 加滿(做多少累計) 有可能是0
                        
                        # 做完當下就要檢查L4到幾趴了
                        if (project_level_days.loc[name, 'L4'] + finish_pm) > L4_lower_limit:  # 超過25%
                            finish_pm = L4_lower_limit - project_level_days.loc[name, 'L4']  # 只能做25%
                        project_level_days.loc[name, 'L4'] += finish_pm  # 加入做多少天df
                        project_and_module_finish_PM.loc['PM', name] += finish_pm
                        project_and_module_remain.loc['PM', name] -= finish_pm # 做完歸0
                        
                        # 算成本
                        if finish_pm > 22:  # 有加班
                            finish_pm_norm = 22
                            finish_pm_add = finish_pm - 22
                        else:  # 沒加班
                            finish_pm_norm = finish_pm
                            finish_pm_add = 0  # 0
                        project_level_days_norm.loc[name, 'L4'] += finish_pm_norm  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, 'L4'] += finish_pm_add    # 加班
                        
                        ##############################################
                        # 做專長
                        finish_prof = (param_PM - finish_pm)  # 做滿30
                        
                        # 做完當下就要檢查L4到幾趴了
                        if (project_level_days.loc[name, 'L4'] + finish_prof) > L4_lower_limit:  # 超過25%
                            finish_prof = L4_lower_limit - project_level_days.loc[name, 'L4']  # 只能做25%
                        
                        
                        # 做完當下還要檢查是否超過專長模組
                        if project_and_module_remain.loc[prof, name] < finish_prof:  # 已超過module總和
                            finish_prof = project_and_module_remain.loc[prof, name]
                        
                        project_level_days.loc[name, 'L4'] += finish_prof  # 加入做多少天df
                        project_and_module_finish_PM.loc[prof, name] += finish_prof
                        project_and_module_remain.loc[prof, name] -= finish_prof  # 更新剩餘人天
                        
                        # 算成本
                        if finish_pm > 22:  # 做pm就有加班
                            finish_prof_norm = 0
                            finish_prof_add = finish_prof
                        else:  # pm沒加班
                            if finish_pm + finish_prof > 22:
                                finish_prof_norm = 22 - finish_pm
                                finish_prof_add = finish_pm + finish_prof - 22
                            else:
                                finish_prof_norm = finish_prof
                                finish_prof_add = 0        
                        project_level_days_norm.loc[name, 'L4'] += finish_prof_norm  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, 'L4'] += finish_prof_add    # 加班
                        
                        
                    
                    else:  # 不能做專長
                        project_level_days.loc[name, 'L4'] += param_PM  # 加入做多少天df
                        project_and_module_finish_PM.loc['PM', name] += param_PM
                        project_and_module_remain.loc['PM', name] -= param_PM  # 扣掉已做
                        
                        # 算成本
                        finish_pm_norm = 22
                        finish_pm_add = param_PM - 22
                        project_level_days_norm.loc[name, 'L4'] += finish_pm_norm  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, 'L4'] += finish_pm_add    # 加班
                # print(i, project_and_module_finish_PM)
                # 統計該月finish
                project_and_module_finish_PM_dict[i] = project_and_module_finish_PM.copy()

            
            ################################第二步###############################################
            level_and_module_remain2 = level_and_module.copy()  # 新的人力(數)   
            # 分非PM--跨專案(要按比例分攤至各專案)
            # L4拆出來先分到都剛好25%、各模組都要有L4
            for i in range(earlist_month, latest_month + 1): # 7-12
                # 更新到最新可用的人力->每月都要先更新(專案結束要把人力丟出來)
                for key in project_time_dict:  # 專案一~六
                    if project_time_dict[key][1] + 1 == i:  # 上個月結束，人力(不能跨專案)交出來
                        level_and_module_remain2.loc[PMs_each_project[key], 'L4'] += 1
                
                # 更新PM人數而已
                for name in month_project[i]:  # 該月進行的專案
                    prof = PMs_each_project[name]  # 該專案的PM專長
                    # 只有第一個月要扣掉人力(不能跨專案)
                    if project_time_dict[name][0] == i:  # 專案的第一個月
                        level_and_module_remain2.loc[prof, 'L4'] -= 1  # 用掉人力(PM)
                
                
                for md in project_and_module_remain.index:  # PP MM SD CO FI
                    if md == 'PM':
                        continue
                   
                    finish = level_and_module_remain2.loc[md, 'L4'] * 22  # 做滿22(總數 不能動) 分給該月專案
                    free_L4 = level_and_module_remain2.loc[md, 'L4'] * 22  # 滾動調整
               
                    # 先算remain的total(分攤比例分母--鎖定直到該月該模組都分完)
                    total_L4 = 0
                    for name in month_project[i]:  # 該月進行的專案
                        total_L4 += project_and_module_remain.loc[md, name]
                    
                    # 開始按比例分下去(先分22不加班)
                    for name in month_project[i]:  # 該月進行的專案
                        L4_lower_limit = project.loc["project_days", name] * 0.25  # 該專案L4下限
                        
                        finish_each_project = finish * project_and_module_remain.loc[md, name] / total_L4
                        if (project_level_days.loc[name, 'L4'] + finish_each_project) > L4_lower_limit:  # 超過25%
                            finish_each_project = L4_lower_limit - project_level_days.loc[name, 'L4']  # 做25%即可
                            
                            
                        if finish_each_project > project_and_module_remain.loc[md, name]:  # 超過module總和
                            finish_each_project = project_and_module_remain.loc[md, name]
                    
                        project_and_module_finish_L4.loc[md, name] += finish_each_project  # finish
                        project_level_days.loc[name, 'L4'] += finish_each_project  # 加入做多少天df (分攤後)
                        project_and_module_remain.loc[md, name] -= finish_each_project  # remain =0
                        free_L4 -= finish_each_project
                        
                        # 算成本
                        project_level_days_norm.loc[name, 'L4'] += finish_each_project  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, 'L4'] += 0    # 加班
                    
                    # 分完一個模組 看該模組專長是否有閒置
                    if free_L4 > 0:
                        level_idle.loc[len(level_idle)] = ['L4_n', i, md, free_L4]
                
                # 統計該月finish
                project_and_module_finish_L4_dict[i] = project_and_module_finish_L4.copy()
                  


            # print(level_and_module_remain)
            # print(project_level_days)  # 用來算成本
            # print(project_and_module_finish_PM)
            # print(project_and_module_finish_L4)  # L4 L3 L2 L1
            # print(project_and_module_finish_L3)
            # print(project_and_module_finish_L2)
            # print(project_and_module_finish_L1)
            # print(project_and_module_remain)

            
            ################################第三步##########################################
            # L4做到25%即可 沒做滿25%要補滿
            L4_add_dict = {}  # 各專案L4要砍(-)或加(+)多少
            for pj in project_and_module_remain:  # 專案一二三
                L4_lower_limit = project.loc["project_days", pj] * 0.25  # L4各專案下限
                if project_level_days.loc[pj, 'L4'] != L4_lower_limit:  # 與25%不符
                    add_l4 = L4_lower_limit - project_level_days.loc[pj, 'L4']  # 砍/加多少
                    L4_add_dict[pj] = add_l4
            # print(L4_add_dict)

            level_and_module_remain3 = level_and_module.copy()  # 新的人力(數)   
            # 補L4沒到25%的
            for i in range(earlist_month, latest_month + 1): # 7-12
                # 更新到最新可用的人力->每月都要先更新(專案結束要把人力丟出來)
                for key in project_time_dict:  # 專案一~六
                    if project_time_dict[key][1] + 1 == i:  # 上個月結束，人力(不能跨專案)交出來
                        level_and_module_remain3.loc[PMs_each_project[key], 'L4'] += 1
                
                # 更新PM人數而已
                for name in month_project[i]:  # 該月進行的專案
                    prof = PMs_each_project[name]  # 該專案的PM專長
                    # 只有第一個月要扣掉人力(不能跨專案)
                    if project_time_dict[name][0] == i:  # 專案的第一個月
                        level_and_module_remain3.loc[prof, 'L4'] -= 1  # 用掉人力(PM)
                
                
                for md in project_and_module_remain.index:  # PP MM SD CO FI
                    if md == 'PM':
                        continue
             
                    finish = level_and_module_remain3.loc[md, 'L4'] * 8  # 做滿8(總數加班) 分給該月專案
                    free_L4a = level_and_module_remain3.loc[md, 'L4'] * 8  # 滾動調整
              
                    # 先算remain的total(分攤比例分母:用差多少25%人力--鎖定直到該月該模組都分完)
                    total_L4 = 0
                    for name in month_project[i]:  # 該月進行的專案
                        if name not in L4_add_dict:  # 已達25%的專案不用參與分攤
                            continue
                        total_L4 += L4_add_dict[name]
                    
                    # 開始按比例分下去(分加班8)
                    for name in month_project[i]:  # 該月進行的專案
                        if name not in L4_add_dict:  # 已達25%的專案不用參與分攤
                            continue
                            
                        L4_lower_limit = project.loc["project_days", name] * 0.25  # 該專案L4下限
                        
                        finish_each_project = finish * L4_add_dict[name] / total_L4
                        if (project_level_days.loc[name, 'L4'] + finish_each_project) > L4_lower_limit:  # 超過25%
                            finish_each_project = L4_lower_limit - project_level_days.loc[name, 'L4']  # 做25%即可
                            
                        if finish_each_project > project_and_module_remain.loc[md, name]:  # 超過module總和
                            finish_each_project = project_and_module_remain.loc[md, name]
                    
                        project_and_module_finish_L4.loc[md, name] += finish_each_project  # finish
                        project_and_module_finish_L4_dict[i].loc[md, name] += finish_each_project
                        project_level_days.loc[name, 'L4'] += finish_each_project  # 加入做多少天df (分攤後)
                        project_and_module_remain.loc[md, name] -= finish_each_project  # remain =0
                        free_L4a -= finish_each_project
                        
                        # 算成本
                        project_level_days_norm.loc[name, 'L4'] += 0  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, 'L4'] += finish_each_project    # 加班
                    
                    # 分完一個模組 看該模組專長是否有閒置
                    if free_L4a > 0:
                        level_idle.loc[len(level_idle)] = ['L4_a', i, md, free_L4a]
                    
                   
                     

            # print(level_and_module_remain)
            # print(project_level_days)  # 用來算成本
            # print(project_and_module_finish_PM)
            # print(project_and_module_finish_L4)  # L4 L3 L2 L1
            # print(project_and_module_finish_L3)
            # print(project_and_module_finish_L2)
            # print(project_and_module_finish_L1)
            # print(project_and_module_remain)            

            # double check--L4是否都剛好25%
            L4_add_dict = {}  # 各專案L4要砍(-)或加(+)多少
            for pj in project_and_module_remain:  # 專案一二三
                L4_lower_limit = project.loc["project_days", pj] * 0.25  # L4各專案下限
                if project_level_days.loc[pj, 'L4'] != L4_lower_limit:  # 與25%不符
                    add_l4 = L4_lower_limit - project_level_days.loc[pj, 'L4']  # 砍/加多少
                    L4_add_dict[pj] = add_l4
            # print(L4_add_dict)         



            ########################第四步#################################################
         
            level_and_module_remain4 = level_and_module.copy()  # 新的人力(數)   
            # 分L1 L2 L3--正常+加班跨專案(要按比例分攤至各專案)
            for hour in [22, 8]:
                for i in range(earlist_month, latest_month + 1): # 7-12
                    # 更新到最新可用的人力->每月都要先更新(專案結束要把人力丟出來)
                    for key in project_time_dict:  # 專案一~六
                        if project_time_dict[key][1] + 1 == i:  # 上個月結束，人力(不能跨專案)交出來
                            level_and_module_remain4.loc[PMs_each_project[key], 'L4'] += 1
                    
                    # 更新PM人數而已
                    for name in month_project[i]:  # 該月進行的專案
                        prof = PMs_each_project[name]  # 該專案的PM專長
                        # 只有第一個月要扣掉人力(不能跨專案)
                        if project_time_dict[name][0] == i:  # 專案的第一個月
                            level_and_module_remain4.loc[prof, 'L4'] -= 1  # 用掉人力(PM)
                    
                    
                    for md in project_and_module_remain.index:  # PP MM SD CO FI
                        if md == 'PM':
                            continue
                        
                        # 先算remain的total(分攤比例分母--鎖定直到該月該模組都分完)
                        total_a = 0
                        urgent_project = {}  # 此月結束的專案：該模組剩餘人天
                        other_project = {}  # 非此月結束的專案：該模組剩餘人天
                        for name in month_project[i]:  # 該月進行的專案
                            if project_time_dict[name][1] == i:  # 這個月結束的專案優先做完
                                urgent_project[name] = project_and_module_remain.loc[md, name]
                            else:
                                other_project[name] = project_and_module_remain.loc[md, name]
                        
                        # 1此月結束的專案 剩餘人天大的先做
                        urgent_project_inv = {}  # key是remains
                        for key in urgent_project:
                            if urgent_project[key] == 0:
                                continue
                            if urgent_project[key] not in urgent_project_inv:
                                urgent_project_inv[urgent_project[key]] = [key]
                            else:
                                urgent_project_inv[urgent_project[key]].append(key)
                        urgent_order_remain = sorted(list(urgent_project_inv.keys()), reverse=True)  # 由大到小的remain
                        
                        
                        # 2非此月結束的專案 用剩餘人天比例分攤 
                        total_a = sum(list(other_project.values()))
                        
                      
                        for lv in ['L1', 'L2', 'L3']:  # 有模組沒有L1也無所謂(按此順序)
                       
                            finish = level_and_module_remain4.loc[md, lv] * hour  # 做滿8(總數) 分給該月專案
                       
                            # 開始按比例分下去(分8加班)#######################################
                            # 先做1 此月結束的專案
                            for remain in urgent_order_remain:  # 大的remain優先抓出來
                                for name in urgent_project_inv[remain]:
                                    finish_each_project = project_and_module_remain.loc[md, name]
                                    
                                    if finish < project_and_module_remain.loc[md, name]:  # 做不完 需要其他level共同努力
                                        finish_each_project = finish
                                    
                                    if lv == 'L1':  # L1還要另外比上限
                                        L1_lower_limit = project.loc["project_days", name] * 0.2  # 該專案L1上限
                                        if (project_level_days.loc[name, lv] + finish_each_project) > L1_lower_limit:  # 超過20%
                                            finish_each_project = L1_lower_limit - project_level_days.loc[name, lv]  # 做20%即可
                                    
                                    # 加入統計表
                                    if lv == 'L1':
                                        project_and_module_finish_L1.loc[md, name] += finish_each_project  # finish
                                    elif lv == 'L2':
                                        project_and_module_finish_L2.loc[md, name] += finish_each_project  # finish
                                    else:
                                        project_and_module_finish_L3.loc[md, name] += finish_each_project  # finish
                                    
                                    project_level_days.loc[name, lv] += finish_each_project  # 加入做多少天df (分攤後)
                                    project_and_module_remain.loc[md, name] -= finish_each_project  # remain =0
                                    finish -= finish_each_project
                                    
                                    # 算成本
                                    if hour == 22:  # 排不加班
                                        project_level_days_norm.loc[name, lv] += finish_each_project  # 加入做多少天df(非加班)
                                    else:  # 8
                                        project_level_days_add.loc[name, lv] += finish_each_project    # 加班
                            
                            # 統計該月finish
                            if hour == 22:
                                if lv == 'L1':
                                    project_and_module_finish_L1_dict[i] = project_and_module_finish_L1.copy()
                                elif lv == 'L2':
                                    project_and_module_finish_L2_dict[i] = project_and_module_finish_L2.copy()
                                elif lv == 'L3':
                                    project_and_module_finish_L3_dict[i] = project_and_module_finish_L3.copy()
                                    
                            if finish <= 0:
                                continue
                            
                            # 2再做其他的
                            freeL123 = finish  # 滾動調整
                            for name in other_project:  # 該月進行的專案
                                if total_a == 0:  # 表示該月該模組已全數做完
                                    continue  # 去抓下個模組
                                
                                finish_each_project = finish * project_and_module_remain.loc[md, name] / total_a
                                if finish_each_project > project_and_module_remain.loc[md, name]:  # 超過module總和
                                    finish_each_project = project_and_module_remain.loc[md, name]
                                
                                if lv == 'L1':  # L1還要另外比上限
                                    L1_lower_limit = project.loc["project_days", name] * 0.2  # 該專案L1上限
                                    if (project_level_days.loc[name, lv] + finish_each_project) > L1_lower_limit:  # 超過20%
                                        finish_each_project = L1_lower_limit - project_level_days.loc[name, lv]  # 做20%即可
                           
                                # 加入統計表
                                if lv == 'L1':
                                    project_and_module_finish_L1.loc[md, name] += finish_each_project  # finish
                                    project_and_module_finish_L1_dict[i].loc[md, name] += finish_each_project
                                elif lv == 'L2':
                                    project_and_module_finish_L2.loc[md, name] += finish_each_project  # finish
                                    project_and_module_finish_L2_dict[i].loc[md, name] += finish_each_project
                                else:
                                    project_and_module_finish_L3.loc[md, name] += finish_each_project  # finish
                                    project_and_module_finish_L3_dict[i].loc[md, name] += finish_each_project
                                
                                project_level_days.loc[name, lv] += finish_each_project  # 加入做多少天df (分攤後)
                                project_and_module_remain.loc[md, name] -= finish_each_project  # remain =0
                                freeL123 -=  finish_each_project
                                
                                # 算成本
                                if hour == 22:  # 排不加班
                                    project_level_days_norm.loc[name, lv] += finish_each_project  # 加入做多少天df(非加班)
                                else:  # 8
                                    project_level_days_add.loc[name, lv] += finish_each_project    # 加班
                            
                            # 分完一個模組 看該模組專長是否有閒置
                            if freeL123 > 0:
                                level_idle.loc[len(level_idle)] = ['L4_n', i, md, freeL123]
                
                            
                 

            # print(level_and_module_remain)
            # print(project_level_days)  # 用來算成本
            # print(project_and_module_finish_PM)
            # print(project_and_module_finish_L4)  # L4 L3 L2 L1
            # print(project_and_module_finish_L3)
            # print(project_and_module_finish_L2)
            # print(project_and_module_finish_L1)
            # print(project_and_module_remain)     

            ##############################第五步#################################################
            # 微調--L4(有浪費)
            # print(level_idle)  # 微調前

            for md in project_and_module_remain.index:
                for pj in project_and_module_remain.columns:
                    if project_and_module_remain.loc[md, pj] == 0:  # 專案做完了
                        continue
                    level_and_module_remain5 = level_and_module.copy()  # 新的人力(數)--L4才會用到
                    for i in range(project_time_dict[pj][0], project_time_dict[pj][1] + 1): # 起始月-結束月
                        
                        #########L4才會用到##############################################
                        # 更新到最新可用的人力->每月都要先更新(專案結束要把人力丟出來)
                        for key in project_time_dict:  # 專案一~六
                            if project_time_dict[key][1] + 1 == i:  # 上個月結束，人力(不能跨專案)交出來
                                level_and_module_remain5.loc[PMs_each_project[key], 'L4'] += 1
                    
                        # 更新PM人數而已
                        for name in month_project[i]:  # 該月進行的專案
                            prof = PMs_each_project[name]  # 該專案的PM專長
                            # 只有第一個月要扣掉人力(不能跨專案)
                            if project_time_dict[name][0] == i:  # 專案的第一個月
                                level_and_module_remain5.loc[prof, 'L4'] -= 1  # 用掉人力(PM)
                        #############################################################################
                        # 因L1已達20%上限 或沒有人力了(理論L2 L3也應幾乎沒有) 所以排L4>25%
                        finish_each_project = 0
                        finish_each_project_norm = 0
                        finish_each_project_add = 0
                        for row in range(len(level_idle)):  # 一row一row抓來看
                            if (level_idle.loc[row, 'Level'] == 'L4_n' and level_idle.loc[row, 'Month'] == i and level_idle.loc[row, 'Module'] == md) or \
                            (level_idle.loc[row, 'Level'] == 'L4_a' and level_idle.loc[row, 'Month'] == i and level_idle.loc[row, 'Module'] == md):
                                finish_each_project += level_idle.loc[row, 'Idle']
                                if level_idle.loc[row, 'Level'] == 'L4_n' and level_idle.loc[row, 'Month'] == i and level_idle.loc[row, 'Module'] == md:
                                    finish_each_project_norm += level_idle.loc[row, 'Idle']  # 正常
                                else:
                                    finish_each_project_add += level_idle.loc[row, 'Idle']  # 非加班
                                
               
                        if finish_each_project > project_and_module_remain.loc[md, pj]:  # 做完remain就好
                            finish_each_project = project_and_module_remain.loc[md, pj]
                            if finish_each_project_norm > project_and_module_remain.loc[md, pj]:  # 不須加班
                                finish_each_project_norm = finish_each_project
                                finish_each_project_add = 0
                            else:  # 還須加班
                                finish_each_project_add = finish_each_project - finish_each_project_norm
                        
                        project_and_module_finish_L4.loc[md, pj] += finish_each_project  # finish
                        project_and_module_finish_L4_dict[i].loc[md, pj] += finish_each_project
                        project_level_days.loc[pj, lv] += finish_each_project  # 加入做多少天df (分攤後)
                        project_and_module_remain.loc[md, pj] -= finish_each_project  # remain =0
                        
                        # 算成本
                        project_level_days_norm.loc[name, lv] += finish_each_project_norm  # 加入做多少天df(非加班)
                        project_level_days_add.loc[name, lv] += finish_each_project_add    # 加班
                    
                        
                            
                                
            # Final
            # print(level_and_module_remain)
            # print(project_level_days)  # 用來算成本
            # print(project_level_days_norm)
            # print(project_level_days_add)
            
            # print(project_and_module_finish_PM)
            # print(project_and_module_finish_L4)  # L4 L3 L2 L1
            # print(project_and_module_finish_L3)
            # print(project_and_module_finish_L2)
            # print(project_and_module_finish_L1)
            # print(project_and_module_remain) 
            # print(level_idle)  # 閒置產能--微調時優先使用

            
            
            ##########################第六步####################################################
            # 檢驗是否有做不完的專案->換下一種PM排法 有做不完就不要做下面
            if (project_and_module_remain.sum() == 0).all():
                #print(permutations1[p])
                feasibility_num += 1  # 為可行解
                
                # 可行解才需去算成本
                # 找總成本最小者
                total_cost = 0
                project_level_days_cost = pd.DataFrame(0, index=project_and_module_remain.columns, columns=level.columns)
                
                
                for pj in project_level_days_cost.index:
                    for lv in project_level_days_cost.columns:
                        pj_lv_norm_cost = round(project_level_days_norm.loc[pj, lv] * level.loc['normal_wage', lv] * 30) # 非加班成本
                        pj_lv_add_cost = round(project_level_days_add.loc[pj, lv] * level.loc['overtime_wage', lv] * 30)  # 加班成本
                        project_level_days_cost.loc[pj, lv] = pj_lv_norm_cost + pj_lv_add_cost
                        
                # 水平加總
                project_level_days_cost['Project_Cost'] = project_level_days_cost.sum(axis=1)
                
                
                if feasibility_num == 1:  # 第一種要強制更新
                    total_cost = project_level_days_cost['Project_Cost'].sum()
                    # print(total_cost)
                    # print(project_level_days_norm)
                    # print(project_level_days_add)
                    # print(project_level_days_cost)
                
                if project_level_days_cost['Project_Cost'].sum() < total_cost:
                    total_cost = project_level_days_cost['Project_Cost'].sum()  # 更新最小成本
                    most_saving_schedule = []  # 清空
                    most_saving_schedule.append(permutations1[p])  # 加上去
           
                elif project_level_days_cost['Project_Cost'].sum() == total_cost:
                    most_saving_schedule.append(permutations1[p])  # 直接加上去
                    
                
                
            if p == len(permutations1) - 1:  # 表示240種都跑完了
                if feasibility_num == 0:  # 沒有可行解
                    # print('這段期間做不完這麼多專案')
                    break
                elif len(most_saving_schedule) != 0 and best_solution == False:  # 第一次輸出
                    return most_saving_schedule  # 有找到最佳解
                    
                elif best_solution == True:  # 第二次輸出
                    # 新增column 表示各專案報價
                    project_level_days_cost['Project_Quote'] = None
                    rev_dict = {}
                    for pj in project_level_days_cost.index:
                        gp_rate = project.loc['gross_profit_rate', pj]
                        project_level_days_cost.loc[pj, 'Project_Quote'] = project_level_days_cost.loc[pj, 'Project_Cost'] * (1 + gp_rate)
                  
                    return project_level_days, project_level_days_cost, project_and_module_finish_PM_dict, \
                    project_and_module_finish_L4_dict, project_and_module_finish_L3_dict, project_and_module_finish_L2_dict, \
                    project_and_module_finish_L1_dict

    # print(feasibility_num)
    # print(most_saving_schedule)

    #################################################################################  
    # 呼叫函數--1先找出最佳解
    #most_saving_schedule = find_best_schedule(project, level, level_and_module, project_and_module, project_time_dict, earlist_month, latest_month, permutations1, choose_round, month_project, best_solution)    
    most_saving_schedule = [('MM', 'PP', 'SD', 'CO', 'FI', 'MM')]
    #print(most_saving_schedule)  # [('MM', 'PP', 'SD', 'CO', 'FI', 'MM')]
    if most_saving_schedule == None:  # 表示沒有解
        PMs_best_schedule_df = 0
        project_level_days = 0
        project_level_days_cost = 0
        PM_schedule = 0
        L4_schedule = 0
        L3_schedule = 0
        L2_schedule = 0
        L1_schedule = 0
        return PMs_best_schedule_df, project_level_days, project_level_days_cost, PM_schedule, L4_schedule, L3_schedule, L2_schedule, L1_schedule
    
    # 針對最佳解--寫出對應各專案之PM專長
    PMs_best_schedule = {}  # 專案：PM專長
    for r in range(len(list(choose_round.values()))):
        for e in range(len(list(choose_round.values())[r])):
            name = list(choose_round.values())[r][e]
            PMs_best_schedule[name] = most_saving_schedule[0][e]
    # print(PMs_best_schedule)  # sheet 1--dict
    PMs_best_schedule_df = pd.DataFrame.from_dict(PMs_best_schedule, orient='index', columns=['PM專長'])  # 轉成dataframe
    # print(PMs_best_schedule_df)  # sheet 1--df

    # 呼叫函數--2針對最佳解找出成本
    if len(most_saving_schedule) != 0:  # 表示有最佳解
        best_solution = True
        project_level_days, project_level_days_cost, PM_schedule, L4_schedule, L3_schedule, L2_schedule, L1_schedule = find_best_schedule(project, level, level_and_module, project_and_module, project_time_dict, earlist_month, latest_month, most_saving_schedule, choose_round, month_project, best_solution)
        # print(project_level_days)         # sheet 2
        # print(project_level_days_cost)    # sheet 3
        # print(PM_schedule)  # sheet 4 --dict
        # print(L4_schedule)  # sheet 5 --dict
        # print(L3_schedule)  # sheet 6 --dict
        # print(L2_schedule)  # sheet 7 --dict
        # print(L1_schedule)  # sheet 8 --dict
        return PMs_best_schedule_df, project_level_days, project_level_days_cost, PM_schedule, L4_schedule, L3_schedule, L2_schedule, L1_schedule
    

#####################################################
# 設計下載介面
from tkinter import Tk, Button, filedialog

def download_excel_financial_data(PMs_best_schedule_df, project_level_days, project_level_days_cost):  # sheet1,2,3
    # 輸出financial_data 因此excel有3個sheet 與其他excel不太相同
    
    # 讓使用者選擇存儲 Excel 檔案的路徑和檔名
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Files", "*.xlsx")])

    # 如果使用者選擇了路徑和檔名，則進行 Excel 輸出
    if filepath:
        # 使用 pandas 將 DataFrame 1 輸出為 Excel 檔案
        writer1 = pd.ExcelWriter(filepath)
        PMs_best_schedule_df.to_excel(writer1, sheet_name='PMs_profesional', index=True)  # 1
        project_level_days.to_excel(writer1, sheet_name='project_level_days', index=True)  # 2
        project_level_days_cost.to_excel(writer1, sheet_name='project_level_days_cost', index=True)  # 3
        writer1.close()  # 儲存
        print("Excel 檔案 financial_data 已成功輸出至:", filepath)
        
        
def download_excel_scheduling_data(excel_name):  # sheet4-8
    # 輸出scheduling_data 都有6個sheet
    
    # 讓使用者選擇存儲 Excel 檔案的路徑和檔名
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Files", "*.xlsx")])

    # 如果使用者選擇了路徑和檔名，則進行 Excel 輸出
    if filepath:
        # 使用 pandas 將 DataFrame 1 輸出為 Excel 檔案
        writer2 = pd.ExcelWriter(filepath)
        for key, value in excel_name.items():
            sheet_name = str(key) + '月'
            value.to_excel(writer2, sheet_name=sheet_name, index=True)
        writer2.close()  # 儲存
        print("Excel 檔案 financial_data 已成功輸出至:", filepath)
        

# 創建主視窗
window = Tk()
window.title("Excel檔案上傳")
window.geometry("300x300")

# 儲存使用者輸入的參數
param_value = None

# 開啟檔案按鈕
open_button = Button(window, text="上傳檔案", command=open_file)
open_button.pack()

# 開始主迴圈
window.mainloop()