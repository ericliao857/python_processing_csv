import os
import pandas as pd
import datetime
import random

# 讀取資料夾內的.csv檔 判斷檔名後進行處理
# 需要把檔案都放在一個資料，並要限制檔名！

## 讀取Path(路徑)內是.csv的檔案
def file_name(file_dir):
    list=[]
    for file in os.listdir(file_dir):
        if os.path.splitext(file)[1] == '.csv':
            list.append(file)
    return list

## 取得系統當前時間
def getNowDate():
    ISOTIMEFORMAT = "%Y%m%d"
    nowDate = datetime.datetime.now().strftime(ISOTIMEFORMAT)
    return nowDate

## 亂數
def randomInt():
    return random.randint(0,999)

## 各ID的處理
def 匯出ID表(df):

    # 設定欄位名稱
    df.columns = ['ID','Day','Hour','Count','remove']

    # 刪除多的事件計數欄位
    df.drop(labels=['remove'], axis='columns', inplace=True)

    # 刪除第0欄 總數的欄位
    df.drop(df.index[0],axis=0,inplace=True)

    # # 總行數
    # print(len(df))
    # 文件的欄位名稱
    print(df.columns)
    # 樞紐分析表 (ID次數表)
    pivot_df = df.pivot_table(index=['ID'],values = 'Count', aggfunc='sum')
    pivot_df = sortTable(pivot_df, 'Count')
    # print(pivot_df.head(5))

    df= sortTable(df,'Count')
    # 總表
    Summary(df)
    # 平日
    Weekday(df)
    # 假日
    Holiday(df)
    # ID次數表
    pivot_df.to_excel(outputPath+fileName+'_ID次數表_'+str(randomInt())+'.xlsx', header=True, index=True)
    # pivot_df.to_csv(outputPath+fileName+'_次數表_'+str(randomInt())+'.csv', header=True, index=True)

## 各村里
def 匯出村里表(df):

    # 設定欄位名稱
    df.columns = ['City',"Town",'Village','Day','Hour','Count','remove']

    # 刪除多的事件計數欄位
    df.drop(labels=['remove'], axis='columns', inplace=True)

    # 刪除第0欄 總數的欄位
    df.drop(df.index[0],axis=0,inplace=True)

    # 文件的欄位名稱
    print(df.columns)

    # 用mergeCityTown 產出City+Twon合併表
    newTable_CityTown = mergeCityTown(df)

    # 用mergeCityTownVillage 產出City+Twon+Village合併表
    newTable_CityTownVillage = mergeCityTownVillage(df)

    # 產出樞紐分析表 (index = 欄 , values = 值 , aggfunc = 加總.平均之類的)

    # 產出City+Town的樞紐分析表
    pivot_CityTown = newTable_CityTown.pivot_table(index=['City_Town'],values = 'Count', aggfunc='sum')
    pivot_CityTown = sortTable(pivot_CityTown,'Count') 

    # 產出City+Town+Village的樞紐分析表
    pivot_village = newTable_CityTownVillage.pivot_table(index=['City_Town_Village'],values = 'Count', aggfunc='sum')
    pivot_village= sortTable(pivot_village,'Count') 
    
    df = sortTable(df,'Count') 
    # 總表
    Summary(df)
    # 平日
    Weekday(df)
    # 假日
    Holiday(df)
    # 縣市鄉鎮次數
    # pivot_CityTown.to_csv(outputPath+fileName+'_縣市鄉鎮次數_'+str(randomInt())+'.csv', header=True, index=True)
    pivot_CityTown.to_excel(outputPath+fileName+'_縣市鄉鎮次數_'+str(randomInt())+'.xlsx', header=True, index=True)
    # 縣市鄉鎮村里次數
    # pivot_village.to_csv(outputPath+fileName+'_縣市鄉鎮村里次數_'+str(randomInt())+'.csv', header=True, index=True)
    pivot_village.to_excel(outputPath+fileName+'_縣市鄉鎮村里次數_'+str(randomInt())+'.xlsx', header=True, index=True)

## 處理排序
def sortTable(sort_df,fieldName):
    sort_df.sort_values(by=[fieldName],ascending=False,inplace=True)
    return sort_df

## 處理CityTown合併
def mergeCityTown(df):
    # 產出City+Twon合併表
    temp = df.City + df.Town
    # 轉成DataFrame格式
    city_town = pd.DataFrame(temp)
    # 設定欄位名稱
    city_town.columns = ['City_Town']
    # print(city_town.head(3))
    newTable = city_town.merge(df,left_index=True, right_index=True)
    return newTable

## 處理CityTownVillage合併
def mergeCityTownVillage(df):
    # 產出City+Twon+Village合併表
    temp = df.City + df.Town + df.Village
    # 轉成DataFrame格式
    city_town_village =pd.DataFrame(temp)
    # 設定欄位名稱
    city_town_village.columns = ['City_Town_Village']
    # print(city_town_village.head(3))
    newTable_village = city_town_village.merge(df,left_index=True, right_index=True)
    return newTable_village

## 匯出總表
def Summary(summary_df):
    # print('匯出總表: 開始！')
    summary_df.to_excel(outputPath+fileName+'_彙總表_'+str(randomInt())+'.xlsx', header=True, index=False)
    # print('匯出總表: 結束！')

## 匯出假日的表
def Weekday(day_df):
    # print('匯出假日表: 開始！')
    weekday_df = day_df.query('Day == 6 or Day == 7')
    # print(weekday_df.head(3))
    weekday_df.to_excel(outputPath+fileName+'_彙總表(假日)_'+str(randomInt())+'.xlsx', header=True, index=False)
    # print('匯出假日表: 完成！')

## 匯出假日的表
def Holiday(day_df):
    # print('匯出平日表: 開始！')
    holiday_df = day_df.query('Day != 6 and Day != 7')
    # print(holiday_df.head(3))
    holiday_df.to_excel(outputPath+fileName+'_彙總表(平日)_'+str(randomInt())+'.xlsx', header=True, index=False)
    # print('匯出平日表: 完成！')

## 判斷是公廁還是定檢站
def judgmentType(fileNameStr):
    if '定檢站' in fileNameStr:
        return '定檢站'
    elif '公廁' in fileNameStr:
        return '公廁'
    else :
        print('ERROR:'+ fileNameStr +' 檔名找不到公廁及定檢站')
        print()
        return ('找不到公廁及定檢站')

## 判斷是ID還是縣市鄉鎮村里
def judgmentMethod(fileNameStr):
    if 'ID' in fileNameStr:
        return 'ID'
    elif '村里' in fileNameStr:
        return '村里'
    else :
        print('ERROR:'+ fileNameStr +' 檔名找不到ID或村里')
        print()
        return ('找不到ID或村里')

## 判斷是使用ID還是村里的匯出方法
def outputMethod(method):
    if 'ID' in method:
        匯出ID表(read_csv)
    elif '村里' in method:
        匯出村里表(read_csv)

## 判斷路徑正確性
def checkPath(inputPathStr):
    return os.path.exists(inputPathStr)

## 處理路徑
def processPath(path):
    # 把 '\' 轉為 '/' 
    return path.replace('\\','/')+'/'

## 新增匯出資料夾
def mkdir(path):
    #判斷結果
    if not checkPath(path):
        #如果不存在，則建立新目錄
        os.makedirs(path)
        print('-----建立成功-----')

    else:
        #如果目錄已存在，則不建立，提示目錄已存在
        print(path+' <--目錄已存在')

### 主程序 START
# 取得日期
nowData = getNowDate()

# 設定檔案 path (也可以用input，但是我懶 )
inputPath = input('請輸入資料夾路徑: ')

# 找不到資料夾無法進行下一步
count = 0
while not checkPath(inputPath):
    print('找不到資料夾，錯誤三次即跳出程式！')
    inputPath = input('請重新輸入資料夾路徑: ')
    count+=1
    if count >= 2 :
        print('輸入資料夾錯誤三次跳出程式！')
        os.system("pause")
        os._exit(1)
        break
    
# 處理inputPath
path = processPath(inputPath)
print('檔案資料夾: ' + path)

# 用file_name去找路徑內檔案的名稱
wks = file_name(path)

# 輸出路徑
outputPath = path+'output/'
mkdir(outputPath)
print('匯出資料夾: '+ outputPath)
print() 
# 判斷檔案的處理方式，及取得檔案名稱
if len(wks) < 1 :
    print('找不到檔案！')
for i in range(len(wks)):
    # 讀檔
    read_csv = pd.read_csv(path+wks[i],header=6,index_col=False)
    # print(wks[i])
    # 判斷類別 是公廁還是定檢站
    tableType = judgmentType(wks[i])
    # 判斷方法 是ID還是村里
    method = judgmentMethod(wks[i])
    # 判斷檔名是否正常
    if ('找不到' not in tableType and '找不到' not in method):
        # 產出檔案前綴名稱
        fileName = str(getNowDate())+ '_' + tableType + '_' + method
        print('-----開始處理: '+ wks[i]+ '！')
        # print('Type: '+tableType+" Method: "+method+' 匯出檔名:'+fileName)
        outputMethod(method)
        print('-----處理完成: '+ wks[i]+ '！')
        print()
        
os.system("pause")
