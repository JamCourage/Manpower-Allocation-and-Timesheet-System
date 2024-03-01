# Manpower Allocation and Timesheet System
Written by JamCourage                    

**主題類別**            
Scheduling, Operation Research, Optimization Method, Project Management, Cost Management, GUI                      

**分析工具**      
Python(numpy, pandas, datetime, random, tkinter, tkcalendar)

**專案背景**                
與伊雲谷數位科技財務部門、技術部門合作的專案，該公司有一主要業務係協助客戶導入ERP系統，
每一個客戶都視為一個獨立專案，我們必須為公司建置有關專案的**報價** 、**人力排程規劃** 與 **工時表系統** 。          

**專案架構圖**                        
參考：[Manpower_allocation_and_timesheet_explaination.png](Manpower_allocation_and_timesheet_explaination.png)                

**主要目標**      
此專案有以下兩大主軸，各主軸之目標詳見如下：            
(一) 建置專案之排程與報價系統             
	1. 在考慮各式限制式下，降低(最小化)各專案之人力成本，並產出最佳化報價        
	2. 減少閒置產能，將產能做最佳分配         
(二) 建置工時表系統            
	1. 為公司員工建立客製化工時系統
	2. 執行績效差異分析、編製次期預算

**兩大步驟**            
1. 建置專案之排程與報價系統      
	(1) 訂定專案層級、模組層級、員工層級等限制式                 
	(2) 考慮一段期間有多項專案的情況     
	(3) 設計排程演算法               
	(4) 將演算法程式化
	【程式碼】            
	程式碼可參考[scheduling_system.py](scheduling_system)        
	【公司輸入表格】           
	導入排程系統前，公司必須先填的資料包含專案資訊、模組資訊、人力資訊等，
	表格可參考[01import_data.xlsx](01import_data.xlsx)
	【輸出結果】       
	(1) 計算各專案成本與報價結果-> [02financial_data.xlsx](02financial_data.xlsx)          
	(2) 各階層人力排程結果(以level 1人力配置為例)-> [03schedulingL1_data.xlsx](03schedulingL1_data.xlsx)
   
2. 建置工時表系統                    
	(1) 設計客製化工時表      
	(2) 設計提醒填寫工時表功能、進度提醒功能      
	(3) 結合排程系統，執行差異分析       
	(4) 次期預算編製與策略建議                           
   【程式碼】            
   工時系統程式碼可參考[timesheet_system.py](timesheet_system.py)                               
   【工時表範本】              
   可參考[04timesheet_example.xlsx](04timesheet_example.xlsx)                     
   
