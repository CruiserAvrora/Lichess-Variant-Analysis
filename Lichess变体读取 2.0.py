import re
import json
import xlsxwriter

# 读取lichess对局记录并保存主要参数
# 需要保存的参数：对局时间，开局前4步，用时1，用时2，结果1-胜负，结果2-原因，白elo，黑elo，总步数
# 暂时只保存至少5回合的elo和胜负结果
# 按年月保存

# 基本参数

pgn_amount = 105
# 一次读取文件数
elo_range = 25
# 每个分段的elo范围，向上取整
pgn_address = "D:\hordetest\lichess_db_horde_rated_"
# pgn文件位置
save_name = "C:/Users/ad674/Downloads/horde25.xlsx"
# 存储文件位置
player_elo_range = 4000
max_elo = player_elo_range//elo_range
# 选手最大elo
start_year = 2015
start_month = 4
end_year = 2024
end_month = 1
all_pgn_amount = (end_year-start_year+1)*12 + end_month - start_month
month_list = ['01','02','03','04','05','06','07','08','09','10','11','12']
# 起止年月
book = xlsxwriter.Workbook(save_name)
sheet1 = book.add_worksheet('Total')
sheet2 = book.add_worksheet('Ultrabullet')
sheet3 = book.add_worksheet('Bullet')
sheet4 = book.add_worksheet('Blitz')
sheet5 = book.add_worksheet('Rapid')
sheet6 = book.add_worksheet('Classical')
sheet_dict = {1:sheet1,2:sheet2,3:sheet3,4:sheet4,5:sheet5,6:sheet6}
# 创建表
count = 0
# 清除计数，开始写文件

# 这个函数按月读取每个elo分段的总局数、有效局数、黑/白-胜/负/和局数
# 后面重写把黑白归一化得分整合进来
def read_horde(pgn_file,elo_range,max_elo):
    with open(pgn_file,'r') as file:
        # 排序：白胜,白和,白负,黑胜,黑和,黑负
        result_ub,result_bu,result_bl,result_rp,result_cl,result_total = {},{},{},{},{},{}
        for j in range(1,max_elo):
            i = 25 * j
            result_ub[i] = [0,0,0,0,0,0,0,0]
            result_bu[i] = [0,0,0,0,0,0,0,0]
            result_bl[i] = [0,0,0,0,0,0,0,0]
            result_rp[i] = [0,0,0,0,0,0,0,0]
            result_cl[i] = [0,0,0,0,0,0,0,0]
            result_total[i] = [0,0,0,0,0,0,0,0]
        index_dict = {'ultrabullet':result_ub,'bullet':result_bu,'blitz':result_bl,'rapid':result_rp,'classic':result_cl}
        line = file.readline()
        match_time = re.match(r'.*(....)-(..)',pgn_file)
        yyyymm = match_time.group(1) + match_time.group(2)
        print(yyyymm)
        count = 1
        gamecount = 0
        invalidgamecount = 0
        previouscount = 0
        while line:
            if re.match(r'.Event.*',line):
                if invalidgamecount == previouscount:
                    if gamecount != 0:
                        print(time,result,white_elo,black_elo)
                        win_rate = 1/(1 + 10**((black_elo-white_elo)/400))
                        if result == '1-0':
                            result_total[white_elo][0] += 1
                            result_total[black_elo][5] += 1
                            result_total[white_elo][6] += 0.5/win_rate
                            index_dict[time][white_elo][0] += 1
                            index_dict[time][black_elo][5] += 1
                            index_dict[time][white_elo][6] += 0.5/win_rate
                        elif result == '0-1':
                            result_total[white_elo][2] += 1
                            result_total[black_elo][3] += 1
                            result_total[black_elo][7] += 0.5/(1-win_rate)
                            index_dict[time][white_elo][2] += 1
                            index_dict[time][black_elo][3] += 1
                            index_dict[time][black_elo][7] += 0.5/win_rate
                        elif result == '1/2-1/2':
                            result_total[white_elo][1] += 1
                            result_total[black_elo][4] += 1
                            result_total[white_elo][6] += 0.25/win_rate
                            result_total[black_elo][7] += 0.25/(1-win_rate)
                            index_dict[time][white_elo][1] += 1
                            index_dict[time][black_elo][4] += 1
                            index_dict[time][white_elo][6] += 0.25/win_rate
                            index_dict[time][black_elo][7] += 0.25/(1-win_rate)
                else:
                    previouscount = invalidgamecount
                gamecount += 1
                print('gamecount ',gamecount,' invalidgamecount ',invalidgamecount)
            elif invalidgamecount > previouscount:
                pass
            elif re.match(r'.Result "(.*)"',line):
                if re.match(r'.Result "(.*)"',line).group(1) in ['1-0','0-1','1/2-1/2','0-0']:
                    result = re.match(r'.Result "(.*)"',line).group(1)
                else:
                    invalidgamecount += 1
                    pass
            elif re.match(r'.WhiteElo "(.*)"',line):
                try:
                    white_elo = int(re.findall(r'(\d+)',line)[0])
                    white_elo = ((white_elo+(elo_range//2))//elo_range) * elo_range
                except:
                    invalidgamecount += 1
            elif re.match(r'.BlackElo "(.*)"',line):
                try:
                    black_elo = int(re.findall(r'(\d+)',line)[0])
                    black_elo = ((black_elo+(elo_range//2))//elo_range) * elo_range
                except:
                    invalidgamecount += 1
            elif re.match(r'.TimeControl.*',line):
                game_time = re.findall(r'(\d+)',line)
                if game_time:
                    time_control = float(game_time[0]) + 40*float(game_time[1])
                else:
                    invalidgamecount += 1
                if time_control <= 29:
                    time = 'ultrabullet'
                elif time_control <= 179:
                    time = 'bullet'
                elif time_control <= 479:
                    time = 'blitz'
                elif time_control <= 1499:
                    time = 'rapid'
                else:
                    time = 'classic'
            elif re.match(r'.FEN.*',line):
                if line != '[FEN "rnbqkbnr/pppppppp/8/1PP2PP1/PPPPPPPP/PPPPPPPP/PPPPPPPP/PPPPPPPP w kq - 0 1"]\n':
                    print('invalid FEN')
                    invalidgamecount += 1
            elif re.match(r'1.*',line):
                if not re.search(r'[5...]',line):
                    print('invalid game')
                    invalidgamecount += 1
            line = file.readline()
            count += 1
    file.close()
    return result_total,result_ub,result_bu,result_bl,result_rp,result_cl,gamecount,invalidgamecount

# pyexcel需要的列查找函数，xlsxwriter不需要
def find_column(number):
    excel_index = ('','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
    a = number//676
    b = (number - a*676)//26
    c = (number%26 + 1)
    column = excel_index[a] + excel_index[b] + excel_index[c]
    return column


# 无聊的赋值写文件部分，稍后也写成函数，依赖read_horde
while count <= pgn_amount:
    year = str((count+start_month-1)//12 + start_year)
    month = month_list[(count+start_month-1)%12]
    pgn_file = pgn_address + year + "-" + month + ".pgn"
    result_total,result_ub,result_bu,result_bl,result_rp,result_cl,gamecount,invalidgamecount = read_horde(pgn_file,elo_range,max_elo)
    result_dict = {1:result_total,2:result_ub,3:result_bu,4:result_bl,5:result_rp,6:result_cl}
    for i in range (1,7):  
        for j in range(1,max_elo):
            if count == 0:
                sheet_dict[i].write(j,0,j*25)
            sheet_dict[i].write(0,8*count+1,year + month)
            sheet_dict[i].write(0,8*count+2,gamecount)
            sheet_dict[i].write(0,8*count+3,invalidgamecount)
            for k in range (1,9):
                sheet_dict[i].write(j,k+8*count,result_dict[i][j*25][k-1])
    count += 1

for i in range(1,8*count):
    col = find_column(i)
    add_up = '=SUM(' + col + '2:' +  col + str(max_elo) + ')'
    for j in range(1,7):
        sheet_dict[j].write(max_elo,i,add_up)
    if i%8 == 0:
        col2 = find_column(i-5)
        col3 = find_column(i)
        add_up_2 = '=0.5*SUM(' + col2 + str(max_elo+1) + ':' +  col3 + str(max_elo+1) + ')'
        for j in range(1,7):
            sheet_dict[j].write(0,i-1,add_up_2)


book.close()        



'''
data_str = json.dumps(result_total)
with open('C:/Users/ad674/Downloads/stat_result.txt', 'w') as target_file:
	target_file.write(data_str)
target_file.close()
'''

