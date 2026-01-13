"""
创建示例Excel文件用于测试
"""

import pandas as pd


def create_example_data():
    """创建示例Excel文件"""
    # 创建示例数据
    data1 = {
        '产品': ['苹果', '香蕉', '橙子', '葡萄'],
        '数量': [100, 200, 150, 80],
        '单价': [5.5, 3.2, 4.8, 12.0],
        '总价': [550, 640, 720, 960]
    }

    data2 = {
        '姓名': ['张三', '李四', '王五'],
        '部门': ['销售部', '技术部', '人事部'],
        '入职日期': ['2020-01-15', '2019-06-20', '2021-03-10'],
        '工资': [8000, 12000, 7000]
    }

    # 创建DataFrame
    df1 = pd.DataFrame(data1)
    df2 = pd.DataFrame(data2)

    # 写入Excel文件，包含多个工作表
    with pd.ExcelWriter('example.xlsx', engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='产品销售', index=False)
        df2.to_excel(writer, sheet_name='员工信息', index=False)

    print("✓ 已创建示例文件: example.xlsx")
    print("  包含两个工作表: '产品销售' 和 '员工信息'")


if __name__ == "__main__":
    create_example_data()
