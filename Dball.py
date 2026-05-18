#coding: utf-8
'''
随机生成满足条件的100000组双色球红球号码，然后再随机选5组球
'''
from modules import logs
from modules.bbqtool import add_zero
level = 20    #日志级别  0:NOTEST ; 10:DEBUG ; 20:INFO ; 30:WARN ; 40:ERROR ; 50:CRITIAL
name = 'Dball'    #日志文件名
logger = logs.LogUtils(name,level).get_log()

file = "C:\\pythoncode\\vscode\\DoubleBall\\data\\history.xlsx"    #历史记录文件路径

# 读取历史记录文件中的original sheet中的H列红球号码数据,返回一个元组，元组中的每个元素是一个包含红球号码,前五数据的元组

def get_history_red_balls(file):
    import pandas as pd
    df = pd.read_excel(file, sheet_name='original', dtype=str)  # 读取Excel文件中的original sheet，指定数据类型为字符串
    red_balls = []
    firth5 = []
    firth4 = []
    after5 = []
    after4 = []
    for index, row in df.iterrows():
        red_balls.append(row['红球'])
        firth5.append(row['前五'])
        firth4.append(row['前四'])
        after5.append(row['后五'])
        after4.append(row['后四'])
    return red_balls, firth5, firth4, after5, after4

if __name__ == '__main__':
    import random
    logger.info("开始随机生成双色球号码...")

    history_red_balls, history_firth5, history_firth4, history_after5, history_after4 = get_history_red_balls(file)
    logger.info(f"从历史记录文件中获取到{len(history_red_balls)}组历史红球号码")

    # 生成500000组双色球号码
    counter = 0
    red_balls = []
    while counter < 10000:
        red_ball = random.sample(range(1, 34), 6)  # 从1-33中随机选6个红球
        red_ball = add_zero(sorted(red_ball))  # 红球前面补0，使其长度变为2

        red = ''.join(red_ball)  # 将红球号码列表转换为字符串，方便后续核对使用
        firth5 = red[:10]  # 前五个红球号码
        firth4 = red[:8]  # 前四个红球号码
        after5 = red[2:]  # 后五个红球号码
        after4 = red[4:]  # 后四个红球号码

        if (red in history_red_balls) or (firth5 in history_firth5) or (firth4 in history_firth4) or (after5 in history_after5) or (after4 in history_after4):
            continue  # 如果生成的红球号码在历史记录中出现过，则跳过本次循环，继续生成下一组号码
        else:
             counter += 1
             red_balls.append(red_ball)  
    logger.info(f"已生成{counter}组满足条件的双色球红球号码")
    # 随机生成三个红球号码
    selected_red_balls = random.sample(red_balls, 3)

    # 随机生成三个篮球号码
    selected_balls = []
    for i in range(3):
        blue_ball = random.randint(1, 16)  # 从1-16中随机选1个蓝球
        if len(str(blue_ball)) == 1:
            blue_ball = '0' + str(blue_ball)  # 蓝球前面补0，使其长度变为2

        selected_balls.append((selected_red_balls[i], blue_ball))

    # 打印选中的5组球
    for idx, (red_ball, blue_ball) in enumerate(selected_balls):
        #print(f"第{idx + 1}组: 红球: {sorted(red_balls)}, 蓝球: {blue_ball}")
        logger.info(f"第{idx + 1}组: 红球: {red_ball}, 蓝球: {blue_ball} [{''.join(red_ball)}]")

    logger.info("随机生成双色球号码完成！")