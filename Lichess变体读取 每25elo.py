import re
import json
import xlsxwriter

# 读取lichess对局记录并保存主要参数
# 需要保存的参数：对局时间，开局前4步，用时1，用时2，结果1-胜负，结果2-原因，白elo，黑elo，总步数
# 暂时只保存至少5回合的elo和胜负结果
# 按年月保存

def read_horde(pgn_file):
    with open(pgn_file,'r') as file:
        # 排序：白胜,白和,白负,黑胜,黑和,黑负
        result_ub,result_bu,result_bl,result_rp,result_cl,result_total = {},{},{},{},{},{}
        for j in range(1,160):
            i = 25 * j
            result_ub[i] = [0,0,0,0,0,0]
            result_bu[i] = [0,0,0,0,0,0]
            result_bl[i] = [0,0,0,0,0,0]
            result_rp[i] = [0,0,0,0,0,0]
            result_cl[i] = [0,0,0,0,0,0]
            result_total[i] = [0,0,0,0,0,0]
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
                        if result == '1-0':
                            result_total[white_elo][0] += 1
                            result_total[black_elo][5] += 1
                            index_dict[time][white_elo][0] += 1
                            index_dict[time][black_elo][5] += 1
                        elif result == '0-1':
                            result_total[white_elo][2] += 1
                            result_total[black_elo][3] += 1
                            index_dict[time][white_elo][2] += 1
                            index_dict[time][black_elo][3] += 1
                        elif result == '1/2-1/2':
                            result_total[white_elo][1] += 1
                            result_total[black_elo][4] += 1
                            index_dict[time][white_elo][1] += 1
                            index_dict[time][black_elo][4] += 1
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
                if not re.search(r'\?',line):
                    try:
                        white_elo = int(re.findall(r'(\d+)',line)[0])
                        white_elo = ((white_elo+12)//25)* 25
                    except:
                        invalidgamecount += 1
            elif re.match(r'.BlackElo "(.*)"',line):
                if not re.search(r'\?',line):
                    try:
                        black_elo = int(re.findall(r'(\d+)',line)[0])
                        black_elo = ((black_elo+12)//25)* 25
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

def find_column(number):
    excel_index = ('','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
    a = number//676
    b = (number - a*676)//26
    c = (number%26 + 1)
    column = excel_index[a] + excel_index[b] + excel_index[c]
    return column


count = 0
max_elo = 160
month_list = ['01','02','03','04','05','06','07','08','09','10','11','12']
save_name = "C:/Users/ad674/Downloads/horde25.xlsx"
book = xlsxwriter.Workbook(save_name)
sheet1 = book.add_worksheet('Total')
sheet2 = book.add_worksheet('Ultrabullet')
sheet3 = book.add_worksheet('Bullet')
sheet4 = book.add_worksheet('Blitz')
sheet5 = book.add_worksheet('Rapid')
sheet6 = book.add_worksheet('Classical')

# 无聊的赋值写文件部分
while count <= 105:
    year = str((count+3)//12 + 2015)
    month = month_list[(count+3)%12]
    pgn_file = "D:\hordetest\lichess_db_horde_rated_" + year + "-" + month + ".pgn"
    result_total,result_ub,result_bu,result_bl,result_rp,result_cl,gamecount,invalidgamecount = read_horde(pgn_file)
    for i in range(1,max_elo):
        if count == 0:
            sheet1.write(i,0,i*25)
        sheet1.write(0,6*count+1,year + month)
        sheet1.write(0,6*count+2,gamecount)
        sheet1.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet1.write(i,j+6*count,result_total[i*25][j-1])
    for i in range(1,max_elo):
        if count == 0:
            sheet2.write(i,0,i*25)
        sheet2.write(0,6*count+1,year + month)
        sheet2.write(0,6*count+2,gamecount)
        sheet2.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet2.write(i,j+6*count,result_ub[i*25][j-1])
    for i in range(1,max_elo):
        if count == 0:
            sheet3.write(i,0,i*25)
        sheet3.write(0,6*count+1,year + month)
        sheet3.write(0,6*count+2,gamecount)
        sheet3.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet3.write(i,j+6*count,result_bu[i*25][j-1])
    for i in range(1,max_elo):
        if count == 0:
            sheet4.write(i,0,i*25)
        sheet4.write(0,6*count+1,year + month)
        sheet4.write(0,6*count+2,gamecount)
        sheet4.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet4.write(i,j+6*count,result_bl[i*25][j-1])
    for i in range(1,max_elo):
        if count == 0:
            sheet5.write(i,0,i*25)
        sheet5.write(0,6*count+1,year + month)
        sheet5.write(0,6*count+2,gamecount)
        sheet5.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet5.write(i,j+6*count,result_rp[i*25][j-1])
    for i in range(1,max_elo):
        if count == 0:
            sheet6.write(i,0,i*25)
        sheet6.write(0,6*count+1,year + month)
        sheet6.write(0,6*count+2,gamecount)
        sheet6.write(0,6*count+3,invalidgamecount)
        for j in range (1,7):
            sheet6.write(i,j+6*count,result_cl[i*25][j-1])
    count += 1

for i in range (1,6*count-5):
    col = find_column(i)
    add_up = '=SUM(' + col + '2:' +  col + str(max_elo) + ')'
    sheet1.write(max_elo,i,add_up)
    sheet2.write(max_elo,i,add_up)
    sheet3.write(max_elo,i,add_up)
    sheet4.write(max_elo,i,add_up)
    sheet5.write(max_elo,i,add_up)
    sheet6.write(max_elo,i,add_up)
    if i%6 == 0:
        col2 = find_column(i-5)
        col3 = find_column(i)
        add_up_2 = '=0.5*SUM(' + col2 + str(max_elo+1) + ':' +  col3 + str(max_elo+1) + ')'
        sheet1.write(0,i-1,add_up_2)
        sheet2.write(0,i-1,add_up_2)
        sheet3.write(0,i-1,add_up_2)
        sheet4.write(0,i-1,add_up_2)
        sheet5.write(0,i-1,add_up_2)
        sheet6.write(0,i-1,add_up_2)

book.close()        



'''
data_str = json.dumps(result_total)
with open('C:/Users/ad674/Downloads/stat_result.txt', 'w') as target_file:
	target_file.write(data_str)
target_file.close()
'''

