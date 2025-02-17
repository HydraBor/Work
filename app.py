import pandas as pd

# 代练类，包含有效会话数和在线天数
class Dailian:
    def __init__(self, valid_sessions, online_days):
        self.valid_sessions = valid_sessions  # 有效会话数
        self.online_days = online_days  # 在线（天）
        self.contribution = 0  # 代练对号主的贡献（会在计算时更新）
        self.payment = 0  # 代练的支付金额（会在计算时更新）

    def __repr__(self):
        return f"Dailian(valid_sessions={self.valid_sessions}, online_days={self.online_days}, contribution={self.contribution}, payment={self.payment})"

# 号主类，包含代练字典和其他新属性
class Haozhu:
    def __init__(self, name):
        self.name = name  # 号主的名字
        self.dailian_dict = {}  # 代练字典
        # 新增五个属性
        self.actual_valid_sessions = 0  # 实际有效会话数
        self.actual_online_duration = 0  # 实际在线时长
        self.valid_sessions_score = 0  # 有效会话得分
        self.online_duration_score = 0  # 在线时长得分
        self.first_response_score = 0  # 首次响应得分
        
        # 新增两个属性
        self.valid_sessions_contribution = 0  # 有效会话贡献
        self.online_duration_contribution = 0  # 在线时长贡献

    # 添加代练及其属性
    def add_dailian(self, dailian_name, valid_sessions, online_days):
        self.dailian_dict[dailian_name] = Dailian(valid_sessions, online_days)

    # 设置号主的五个新属性，并计算有效会话贡献和在线时长贡献
    def set_scores(self, actual_online_duration, actual_valid_sessions, online_duration_score, 
                   valid_sessions_score, first_response_score):
        self.actual_online_duration = actual_online_duration
        self.actual_valid_sessions = actual_valid_sessions
        self.online_duration_score = online_duration_score
        self.valid_sessions_score = valid_sessions_score
        self.first_response_score = first_response_score
        
        # 计算有效会话贡献和在线时长贡献
        self.valid_sessions_contribution = valid_sessions_score + first_response_score
        self.online_duration_contribution = online_duration_score

    # 计算代练的贡献
    def calculate_dailian_contribution(self):
        for dailian_name, dailian in self.dailian_dict.items():
            # 代练贡献的计算公式
            dailian.contribution = (
                (dailian.valid_sessions / self.actual_valid_sessions) * self.valid_sessions_contribution +
                (dailian.online_days / self.actual_online_duration) * self.online_duration_contribution
            )

    # 计算每个代练应该获得的支付金额
    def calculate_payments(self, k):
        for dailian in self.dailian_dict.values():
            # 计算代练的支付金额
            dailian.payment = dailian.contribution * k * 0.8

    def __repr__(self):
        return f"Haozhu(name={self.name}, dailian_dict={self.dailian_dict}, " \
               f"actual_valid_sessions={self.actual_valid_sessions}, " \
               f"actual_online_duration={self.actual_online_duration}, " \
               f"valid_sessions_score={self.valid_sessions_score}, " \
               f"online_duration_score={self.online_duration_score}, " \
               f"first_response_score={self.first_response_score}, " \
               f"valid_sessions_contribution={self.valid_sessions_contribution}, " \
               f"online_duration_contribution={self.online_duration_contribution})"

# 读取工作详情 CSV 文件
file_path = "工作详情.csv"
df = pd.read_csv(file_path)

# 创建号主字典，键是号主名称，值是Haozhu类实例
haozhu_dict = {}

# 填充号主字典
for index, row in df.iterrows():
    owner_name = row['号主']
    dailian_name = row['代练']
    valid_sessions = row['有效会话数']
    online_days = row['在线（天）']
    
    # 如果号主不存在，创建号主对象
    if owner_name not in haozhu_dict:
        haozhu_dict[owner_name] = Haozhu(owner_name)
    
    # 为号主添加代练
    haozhu_dict[owner_name].add_dailian(dailian_name, valid_sessions, online_days)

# 读取号主数据 CSV 文件
owner_data_path = "号主数据.csv"
owner_df = pd.read_csv(owner_data_path)

# 更新号主的五个新属性（从 "号主数据.csv" 文件读取）
for index, row in owner_df.iterrows():
    owner_name = row['号主']
    actual_online_duration = row['实际在线时长 (天)']
    actual_valid_sessions = row['实际有效会话数']
    online_duration_score = row['在线时长得分']
    valid_sessions_score = row['有效会话得分']
    first_response_score = row['首次响应得分']
    
    # 确保号主已经存在
    if owner_name in haozhu_dict:
        haozhu = haozhu_dict[owner_name]
        # 设置号主的五个新属性
        haozhu.set_scores(actual_online_duration, actual_valid_sessions, online_duration_score, 
                          valid_sessions_score, first_response_score)

# 计算每个代练对号主的贡献
for owner_name, haozhu in haozhu_dict.items():
    haozhu.calculate_dailian_contribution()

# 用户输入当前是4周月还是5周月
weeks_in_month = int(input("请输入当前是几周月（4 或 5）："))

# 根据输入的周数来确定 k 的值
k = 7.5 if weeks_in_month == 4 else 6 if weeks_in_month == 5 else 0
if k == 0:
    print("输入错误，程序终止。")
    exit()

# 计算每个代练的支付金额
for owner_name, haozhu in haozhu_dict.items():
    haozhu.calculate_payments(k)

# 创建一个 DataFrame 来显示结果
payment_data = []
for owner_name, haozhu in haozhu_dict.items():
    for dailian_name, dailian in haozhu.dailian_dict.items():
        payment_data.append({
            "号主": owner_name,
            "代练": dailian_name,
            "代练贡献": dailian.contribution,
            "支付金额 (元)": round(dailian.payment, 2)
        })

# 创建 DataFrame 并显示为表格
payment_df = pd.DataFrame(payment_data)

# 输出结果
print("\n代练支付金额表格：")
print(payment_df)

# 输出到 Excel 文件
payment_df.to_excel("结算.xlsx", index=False)
print("\n支付结果已保存到 '结算.xlsx' 文件中。")