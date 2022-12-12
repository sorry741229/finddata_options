import pandas as pd
import os, sys
from colorama import init, Fore, Back #字體顏色


init(autoreset=True)#字體顏色
#https://clay-atlas.com/blog/2020/03/22/python-%E8%BC%B8%E5%87%BA%E5%AD%97%E4%B8%B2%E5%9C%A8%E7%B5%82%E7%AB%AF%E6%A9%9F%E4%B8%AD%E9%A1%AF%E7%A4%BA%E9%A1%8F%E8%89%B2/

#調整視窗大小
from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT

WindowsSTDOUT = windll.kernel32.GetStdHandle(-11)
dimensions = SMALL_RECT(-10, -10, 120, 40) # (left, top, right, bottom)
# Width = (Right - Left) + 1; Height = (Bottom - Top) + 1
windll.kernel32.SetConsoleWindowInfo(WindowsSTDOUT, True, byref(dimensions))

pd.set_option("display.max_rows", None)    #設定最大能顯示1000rows
pd.set_option("display.max_columns", None) #設定最大能顯示1000columns
pd.set_option('display.width', 1000) # 設置打印寬度 
pd.set_option('display.max_colwidth', 180)
pd.set_option("display.colheader_justify","center") #抬頭對齊用
pd.set_option('display.unicode.ambiguous_as_wide', True) #抬頭對齊用
pd.set_option('display.unicode.east_asian_width', True) #抬頭對齊用


while True :
    print('')
    print(Fore.CYAN +"{:*^100s}".format("歡迎使用群旭科技_CNC表單查詢小幫手"))
    print('')
    try:
        xlsx = input('輸入要查詢Excel路徑: ')
        headlistA = input('Excel的第幾行開始是標籤索引呢? 輸入數字: ')
        headlistA = int(headlistA)
        if isinstance(headlistA, int) : #判斷headlistA是否為int
            pass
        else:
            print('輸入錯誤，請重新輸入')
            continue
        headlistA = headlistA - 1
        while True:
            print('')
            print('輸入要搜尋的 "品項名稱" ','或按' + Fore.CYAN + '【R】','回到上頁或重新輸入查詢的文件路徑: ')
            print('')
            name = input('請輸入 : ')
            if name == 'r' or name == 'R':
                break   
            elif name != 'r' or name != 'R':
                while True:   
                    print('')      
                    print(Fore.YELLOW +"{:=^100s}".format("主要標籤頁面"))
                    print('')
                    print('想使用哪一主要標籤來搜尋' + Fore.BLUE +'【', (name), Fore.BLUE + '】 ','? (文字須與 Excel 完全相同)或按' + Fore.GREEN + '【C】','回到上頁或重新輸入搜尋的品項: ')                      
                    print('')
                    index_input0 = input('請輸入 : ')      
                    if index_input0 == 'c' or index_input0 == 'C':
                        break
                    ws = []
                    list_display0 = [] #希望顯示項目
                    list_display0.append(index_input0)
                    print('')
                    print(Fore.YELLOW +"{:=^100s}".format("主要標籤頁面"))
                    print('')
                    print('')
                    print(Fore.GREEN +'你輸入的搜尋品項的主要搜尋的標籤為', list_display0)
                    print(Fore.GREEN +'你輸入的搜尋品項的主要搜尋的標籤為', list_display0)
                    print(Fore.GREEN +'你輸入的搜尋品項的主要搜尋的標籤為', list_display0)
                    print('')
                    while True:
                        sheet = pd.read_excel(xlsx, sheet_name = None, index_col = 0, header = headlistA)
                        print('')
                        print(Fore.YELLOW +"{:=^100s}".format("次要標籤頁面"))
                        print('')
                        print(Fore.RED +'目前顯示篩選', Fore.CYAN + name, Fore.RED + '索引標籤有', list_display0)
                        print('')
                        print('請增加對其搜尋顯示結果的"索引標籤"，輸入完畢後，')
                        print('')
                        print('請按 '+ Fore.YELLOW + '【Q】','繼續 ，或按' + Fore.CYAN +'【R】','回到上頁或重新輸入主要標籤 ，刪除顯示索引標籤請按' + Fore.RED +'【X】')
                        print('')
                        list_custom0 = input('請輸入 : ')
                        #list_custom1 = '\'' + list_custom0 + '\''
                        print('')
                        print('')
                        if list_custom0 == 'x' or list_custom0 == 'X':
                            print(Fore.RED + '已刪除【', list_display0[-1] + Fore.RED +' 】', '索引標籤')
                            list_display0.pop()
                        elif list_custom0 == 'r' or list_custom0 == 'R':
                            break
                         
                        else:
                            if list_custom0 == 'q' or list_custom0 =='Q':
                                os.system('cls')
                                print('搜尋' + Fore.BLUE +'【', (name), Fore.BLUE + '】','的主要標籤為'+ Fore.CYAN +'【',index_input0 + Fore.CYAN + ' 】')
                                print('')
                                print(Fore.YELLOW +'顯示結果的篩選標籤有', list_display0)
                                print('')
                                print('')
                                print('')
                                print(Fore.CYAN +"{:=^100s}".format("以下為搜尋的結果"))
                                print(Fore.CYAN +"{:=^108s}".format(""))
                                print('')   
                                for line in sheet:
                                    df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header = headlistA)
                                    try:
                                        df = pd.read_excel(xlsx, sheet_name = line, index_col = 0,header = headlistA)
                                        df = df.fillna(' ')
                                        result = df[index_input0].str.contains(name, na=False)#將dataframe轉成文字比對name,得到布林值 
                                        filter_result = df[result]                  #用dataframe操作取出布林後之篩選結果
                                        final_result = filter_result[list_display0] 
                                        if final_result.empty :
                                            continue
                                            #filter_result.set_index(keys = ['品名/規格','廠商品號','單價','備註'],inplace=True) #此為調整索引值順序

                                        else:
                                            list_display0.append(list_custom0)
                                            list_display0.pop()
                                            print(final_result)
                                            print('')
                                            print(Fore.RED +' ----此區塊為', Fore.YELLOW + line.strip(),Fore.RED + '工作表的內容')
                                            print('')
                                            #print('INFO...')
                                            #print(final_result.info(1))
                                            print(Fore.CYAN +"{:=^108s}".format(""))
                                            ws.append(final_result)

                                        #每次使用皆寫入log
                                            final_result.to_excel('//192.168.10.61/f2-cnc報表/other/log.xlsx')
                                        
                                        #每次使用皆寫入log
                                    except :
                                        pass

                                print('')
                                print('')
                                print('')
                                print(Fore.CYAN +"{:=^108s}".format(""))
                                print(Fore.CYAN +"{:=^100s}".format("以上為搜尋的結果"))
                                print('')
                                print(Fore.GREEN + '查詢檔案路徑為: ', xlsx )
                                print(Fore.GREEN + '此檔案共有',len(ws),Fore.GREEN + '個工作表,有找到要查詢的',Fore.CYAN + name, Fore.GREEN + '這個項目')
                                print('')
                                print('')
                                print('')
                                #break

                        
                            else:
                                list_display0.append(list_custom0)
                    



    except :
        print(Fore.RED +'輸入錯誤,請重新輸入')   # 收到錯誤訊息，顯示錯誤
        break
        