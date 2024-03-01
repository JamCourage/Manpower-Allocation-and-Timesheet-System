# Manpower Allocation and Timesheet System
Written by JamCourage                    

**主題類別**            
Scheduling, Operation Research, Project Management                      

**分析工具**      
Python(numpy, pandas, datetime, random, tkinter, tkcalendar)


**專案背景**                
與伊雲谷數位科技財務部門、技術部門合作的專案，該公司有一主要業務係協助客戶導入ERP系統，每一個客戶都視為一個獨立專案，我們必須為公司建置有關專案的報價、人力排程規劃與工時表系統。                 

**主要目標**      
與伊雲谷數位科技財務部門、技術部門合作的專案，此專案有以下兩大主軸：            
(一) 降低各      

**兩大步驟**            
1. [Part 1:蒐集台灣上市櫃公司的英文永續報告書--使用python爬蟲](1.%20Web%20Crawler)      
	(1) 將有揭露英文永續報告書的公司股票代碼存入list中      
	(2) 設定並初始化Chrome     
	(3) 使用for迴圈，迭代每一公司，依序從公開資訊觀測站，下載其英文永續報告書      
	【程式碼】            
	程式碼可參考[web_crawler_for_ESGreports.py](1.%20Web%20Crawler/web_crawler_for_ESGreports.py)，以下載110年永續報告書為例       
   
2. [Part 2:分別計算各永續報告書的語調分數(Tone)--使用FinBERT情緒分類模型 & FinBERT主題分類模型](2.%20FinBERT_calculate%20tone)        
	(1) 整理公司股票代碼：將代碼都存在list中      
	(2) 安裝FinBERT兩大模型、nltk tokenizer      
	(3) 逐一擷取PDF文字(使用pdfplumber)       
	(4) 文字前處理：使用regex套件，保留英文字母大小寫、正常標點符號，其他以空白取代           
	(5) nltk斷句        
	(6) FinBERT：同時檢查不得超過512張量，超過者斷成兩句       
		-(6-1) FinBERT情緒分類模型：將文本分類為**中立** 、**正向** 、**負向**           
		-(6-2) FinBERT主題分類模型：將文本分類為**環境(E)** 、**社會(S)** 、**治理(G)** 、 **非ESG**     	   
	(7) 紀錄結果：使用pandas套件，將結果存成dataframe格式。                
   【程式碼】            
   程式碼可參考：           
   (a) 迭代各公司，計算各公司永續報告書語調分數 [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/JamCourage/Tone-of-Sustainability-Report/blob/main/2.%20FinBERT_calculate%20tone/crawler_finbert.ipynb)                             
   (b) 記錄一家公司每一句的語調與分類  [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/JamCourage/Tone-of-Sustainability-Report/blob/main/2.%20FinBERT_calculate%20tone/crawler_finbert_for_one.ipynb)     
                       
   【輸出結果】              
   107年至110年之各公司永續報告書語調分數(Tone)結果，可參考[Tone_breakdown.xlsx](2.%20FinBERT_calculate%20tone/Tone_breakdown.xlsx)                       
   
