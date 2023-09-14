# 匯入必要的庫
import pandas as pd
from datetime import date, timedelta
import random
import plotly.express as px
import math
import numpy as np

# 讀取工單資料
wo = pd.read_excel('WIP.xlsx')
wo.head(5)

# 定義每個產品類型和機台的加工時間
PROCESS_TIME = {
    ('A', 330): (3600/500, 0),
    ('A', 310): (3600/200, 0),
    ('A', 380): (3600/300, 0),
    ('A', 500): (3600/500, 0),
    ('A', 600): (3600/500, 0),
    ('B', 6000): (3600/180, 0),
    ('B', 20000): (3600/144, 0),
    ('B', 17250): (3600/336, 0),
    ('C', 240): (3600/168, 0),
    ('D', 330): (3600/500, 3600/40),
    ('D', 380): (3600/300, 3600/40),
    ('D', 500): (3600/500, 3600/250)
}

# 計算每個工單的總加工時間
wo['total_time'] = 0

for row in range(len(wo)):
    type = wo.iloc[row]['產品類型']
    machine = wo.iloc[row]['產線']
    wo.at[row, 'total_time'] = wo.iloc[row]['填充需求數量'] * PROCESS_TIME[(machine, type)][0] + wo.iloc[row]['貼標需求數量'] * PROCESS_TIME[(machine, type)][1]

# 定義工單類別
class ORDER():
    def __init__(self, wo_id, wo_type, machine, num_fill, num_sticker, due, total_time):
        self.wo_id = wo_id
        self.wo_type = wo_type
        self.machine = machine
        self.num_fill = num_fill
        self.num_sticker = num_sticker
        self.due = due
        self.total_time = total_time

# 將工單轉換為物件
order_seq = []

for row in range(len(wo)):
    wo_id = wo.iloc[row]['工單編號']
    wo_type = wo.iloc[row]['產品類型']
    machine = wo.iloc[row]['產線']
    num_fill = wo.iloc[row]['填充需求數量']
    num_sticker = wo.iloc[row]['貼標需求數量']
    due = wo.iloc[row]['交期']
    total_time = wo.iloc[row]['total_time']
    order = ORDER(wo_id, wo_type, machine, num_fill, num_sticker, due, total_time)
    order_seq.append(order)

# 定義計算適應度的函數
def fitness_count(order_seq):
    # 初始化資料
    job = []
    start_time = []
    finish_time = []
    resource = []
    machine_state_ABC = {'A': 0, 'B': 0, 'C': 0}  # ABC機台狀況
    machine_state_D = 0  # D機台狀況
    machine_time_ABC = {'A': 0, 'B': 0, 'C': 0}  # ABC機台時間
    machine_time_D = 0  # D機台時間
    all_setup_time = 0  # 時間紀錄
    all_tardiness_time = 0
    previous = None  # 前一單產品類型
    today = date.today()
    today_timestamp = pd.Timestamp(today)

    # 迭代處理每個工單
    for order in order_seq:
        wo_id = order.wo_id
        wo_type = order.wo_type
        machine = order.machine
        num_fill = order.num_fill
        num_sticker = order.num_sticker
        due = order.due

        # 計算填充時間和貼標時間
        fill_time = round(num_fill * PROCESS_TIME[(machine, wo_type)][0], 0)
        stick_time = round(num_sticker * PROCESS_TIME[(machine, wo_type)][1], 0)
        setup_time = 20 * 60

        # 如果前一單和本單的產品類型相同，則整備時間僅為5分鐘
        if previous == wo_type:
            setup_time = 5 * 60

        all_setup_time += setup_time
        total_process_time = fill_time + stick_time + setup_time

        order_start_time = 0
        order_finish_time = 0

        if machine in ['A', 'B', 'C']:
            # 假如前一單是D排單
            if sum(machine_state_ABC.values()) == 0:
                order_start_time = machine_time_D
                order_finish_time = order_start_time + total_process_time
                machine_state_D = 0
                machine_time_ABC[machine] = order_finish_time
                machine_state_ABC[machine] = 1
            # 假如已有一台ABC機群運作
            elif sum(machine_state_ABC.values()) == 1:
                previous_machine = ([current_machine for (current_machine, state) in machine_state_ABC.items() if state == 1])[0]
                if machine == previous_machine:
                    order_start_time = machine_time_ABC[machine]
                    order_finish_time = order_start_time + total_process_time
                    machine_time_ABC[machine] = order_finish_time
                else:
                    order_start_time = machine_time_D
                    order_finish_time = order_start_time + total_process_time
                    machine_state_D = 0
                    machine_time_ABC[machine] = order_finish_time
                    machine_state_ABC[machine] = 1
            else:
                previous_machine = ([current_machine for (current_machine, state) in machine_state_ABC.items() if state == 1])
                if machine in previous_machine:
                    order_start_time = machine_time_ABC[machine]
                    order_finish_time = order_start_time + total_process_time
                    machine_time_ABC[machine] = order_finish_time
                else:
                    ABC_list = [(pre_machine, time) for ((pre_machine, state), time) in
                                (zip(machine_state_ABC.items(), machine_time_ABC.values())) if state == 1]
                    min_element = min(ABC_list, key=lambda x: x[1])
                    order_start_time = min_element[1]
                    order_finish_time = order_start_time + total_process_time
                    machine_state_ABC[min_element[0]] = 0
                    machine_time_ABC[machine] = order_finish_time
                    machine_state_ABC[machine] = 1
        else:
            if machine_state_D == 1:
                order_start_time = machine_time_D
                machine_time_D += total_process_time
                order_finish_time = machine_time_D
            elif sum(machine_state_ABC.values()) > 0:
                ABC_list = [(pre_machine, time) for ((pre_machine, state), time) in
                            (zip(machine_state_ABC.items(), machine_time_ABC.values())) if state == 1]
                min_element = max(ABC_list, key=lambda x: x[1])
                order_start_time = min_element[1]
                order_finish_time = order_start_time + total_process_time
                for history in ABC_list:
                    machine_state_ABC[history[0]] = 0
                machine_time_D = order_finish_time
                machine_state_D = 1
            else:
                order_start_time = machine_time_D
                machine_time_D = total_process_time
                machine_state_D = 1
                order_finish_time = machine_time_D

        previous = wo_type
        finish_date = int(order_finish_time / (3600 * 24))
        tardiness = max((today_timestamp + timedelta(days=finish_date) - due).days, 0)
        all_tardiness_time += tardiness

        job.append(wo_id)
        start_time.append(today_timestamp + timedelta(seconds=order_start_time))
        finish_time.append(today_timestamp + timedelta(seconds=order_finish_time))
        resource.append(machine)

    max_ABC = max(machine_time_ABC.items(), key=lambda x: x[1])[1]
    maxspan = max_ABC
    if max_ABC < machine_time_D:
        maxspan = machine_time_D
    schedule = pd.DataFrame({"工單編號": job, "Start": start_time, "Finish": finish_time, "產線": resource})

    return maxspan, all_setup_time, all_tardiness_time * 3600 * 24, schedule

# 定義個體類別
class Individual:
    def __init__(self, information):
        self.order = information[0]  # 排序順序
        self.fitness = information[1:]  # 適應度值
        self.dominated_solutions = []  # 被支配的解集合
        self.domination_count = 0  # 支配計數器

# 定義NSGA-II演算法
def nsgaii():
    population = []  # 個體群體
    all_pop_fitness = []  # 儲存所有個體的適應度值

    # 隨機生成初始個體
    for _ in range(1000):
        shuffled_order_seq = random.sample(order_seq, len(order_seq))
        maxspan, setuptime, tardiness, sc = fitness_count(shuffled_order_seq)
        all_pop_fitness.append((shuffled_order_seq, maxspan, setuptime, tardiness))

    # 開始進化
    for _ in range(20):  # 世代數
        for fitness in all_pop_fitness:
            individual = Individual(fitness)
            population.append(individual)

        fronts = [[]]  # 非支配排序結果
        for individual in population:
            individual.domination_count = 0
            individual.dominated_solutions = []
            for other in population:
                if all(individual.fitness[i] < other.fitness[i] for i in range(len(individual.fitness))):
                    individual.dominated_solutions.append(other)
                elif all(other.fitness[i] < individual.fitness[i] for i in range(len(individual.fitness))):
                    individual.domination_count += 1
            if individual.domination_count == 0:
                fronts[0].append(individual)

        i = 0
        while len(fronts[i]) > 0:
            next_front = []
            for individual in fronts[i]:
                for other in individual.dominated_solutions:
                    other.domination_count -= 1
                    if other.domination_count == 0:
                        next_front.append(other)
            i += 1
            fronts.append(next_front)

        for front in fronts[0]:
            crowding_list = []
            for other in fronts[0]:
                vector1 = front.fitness
                vector2 = other.fitness
                euclidean_distance = math.sqrt(sum((x - y) ** 2 for x, y in zip(vector1, vector2)))
                crowding_list.append(euclidean_distance)
            point1 = fronts[0][(crowding_list.index(sorted(crowding_list)[1]))]
            point2 = fronts[0][(crowding_list.index(sorted(crowding_list)[2]))]
            
            cross_product = np.cross(point1.fitness, point2.fitness)
            volume = np.linalg.norm(cross_product)

            front.crowd_point = volume
        
        fronts[0].sort(key=lambda x: x.crowd_point, reverse=True)
        mother_order = [front.order for front in fronts[0][:2]]

        new_pops = []
        all_pop_fitness = []
        for _ in range(100):
            random_numbers = sorted(random.sample(range(0, len(order_seq)), 2))
            best_mother = mother_order[0]
            second_mother = mother_order[1]
            find_segement = best_mother[random_numbers[0]:random_numbers[1]+1]
            index_second_seg = sorted([second_mother.index(ele) for ele in find_segement])
            new_segment = best_mother[0:random_numbers[0]] + [second_mother[index] for index in index_second_seg] + best_mother[random_numbers[1]+1:]

            rate = 0.1
            if random.random() < rate:
                random_numbers = sorted(random.sample(range(0, len(order_seq)), 2))
                new_segment[random_numbers[0]], new_segment[random_numbers[1]] = new_segment[random_numbers[1]], new_segment[random_numbers[0]]

            new_pops.append(new_segment)

        for new_pop in new_pops:
            maxspan, setuptime, tardiness, sc = fitness_count(new_pop)
            all_pop_fitness.append((new_pop, maxspan, setuptime, tardiness))

    fronts[0].sort(key=lambda x: x.crowd_point, reverse=True)
    best_ind = fronts[0][0]

    return best_ind.fitness, sc

# 執行NSGA-II演算法並獲取結果
fitness, sc = nsgaii()

# 將時間戳記轉換為日期時間格式
sc["Start"] = pd.to_datetime(sc["Start"])
sc["Finish"] = pd.to_datetime(sc["Finish"])
df = sc.sort_values(by="產線", ascending=False)

# 創建甘特圖
fig = px.timeline(df, 
                  x_start="Start", 
                  x_end="Finish", 
                  y="產線", 
                  color="工單編號", 
                  title="Gantt Chart - by 機台")

# 顯示甘特圖
fig.show()

# 將排程結果合併回原始工單數據並計算遲交時間
df = pd.merge(sc, wo, on='工單編號').drop(columns=['產線_y'])
df['遲繳'] =  df['Finish'] - df['交期']
print(df)
