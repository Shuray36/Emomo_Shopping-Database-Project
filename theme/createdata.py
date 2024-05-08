import pandas as pd
from datetime import datetime, timedelta
import random

# 生成一月的日期
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 1, 31)

date_list = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]

# 生成每天的隨機訂單數量，範圍可以根據實際情況調整
order_quantity_list = [random.randint(1, 10) for _ in range(len(date_list))]

# 生成唯一的Order ID
order_id_list = []
for i, date in enumerate(date_list):
    order_id_list.extend([f'#{i + 1}_{j + 1}' for j in range(order_quantity_list[i])])

# 生成 Total Price 和 Status
total_price_list = [f'${random.randint(50, 200)}.00' for _ in range(len(order_id_list))]
status_list = ['Pending'] * len(order_id_list)

# 創建 data 字典
data = {
    'Order ID': order_id_list,
    'Date': [date.strftime('%b %d, %Y') for date in date_list for _ in range(order_quantity_list[i])],
    'Items': [random.randint(1, 10) for _ in range(len(order_id_list))],
    'Total Price': total_price_list,
    'Status': status_list
}

# 創建 DataFrame
df = pd.DataFrame(data)

# 將 DataFrame 寫入 Excel 檔案
df.to_excel('sales_report.xlsx', index=False)

print("Excel 檔案已經寫入成功！")
