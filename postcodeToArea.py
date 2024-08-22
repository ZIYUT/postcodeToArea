import pandas as pd

# 定义各区的邮编列表
north_area = [3083, 3058, 3058, 3078, 3081, 3084, 3079, 3079, 3083, 3085, 3085, 3070, 3052, 3072, 3073, 3084, 3071,
              3084, 3074]
east_area = [3133, 3152, 3152, 3131, 3150, 3150, 3132, 3170, 3802, 3134, 3136, 3138, 3123, 3111]
east1_area = [3103, 3104, 3124, 3122, 3102, 3101, 3144, 3105, 3975]
east2_area = [3147, 3130, 3129, 3128, 3149, 3151, 3125, 3125, 3109, 3108, 3126, 3107, 3106, 3127,
              3127, 3131, 3754]
south_area = [3143, 3163, 3145, 3162, 3162, 3162, 3148, 3169, 3169, 3168, 3163, 3146, 3166, 3173, 3163, 3174, 3174,
              3168, 3186, 3167, 3166, 3185, 3172, 3171, 3142, 3145, 3144, 3141, 3190, 3141, 3197, 3175, 3188, 3201,
              3180, 3192, 3201]
south1_area = [3165, 3204, 3187, 3187, 3186, 3206, 3206, 3204, 3189]
west_area = [3011, 3029, 3026, 3028, 3024, 3020, 3020, 3030, 3030, 3012, 3027, 3024, 3038, 3019, 3018, 3337]
city_center = [3206, 3057, 3055, 3056, 3054, 3054, 3053, 3068, 3065, 3121, 3103, 3008, 3002, 3031, 3031, 3000, 3207,
               3121, 3205, 3008, 3006, 3183, 3182, 3182, 3003, 3066]

# 根据邮编查找区域
def find_area(postal_code):
    if postal_code in north_area:
        return "北区"
    elif postal_code in east_area:
        return "东区"
    elif postal_code in east1_area:
        return "东1"
    elif postal_code in east2_area:
        return "东2"
    elif postal_code in south_area:
        return "南区"
    elif postal_code in south1_area:
        return "南1"
    elif postal_code in west_area:
        return "西区"
    elif postal_code in city_center:
        return "市中心"
    else:
        possible_areas = []
        if any(str(postal_code).startswith(str(code)[:2]) for code in north_area):
            possible_areas.append("北区")
        if any(str(postal_code).startswith(str(code)[:2]) for code in east_area + east1_area + east2_area):
            possible_areas.append("东区")
        if any(str(postal_code).startswith(str(code)[:2]) for code in south_area + south1_area):
            possible_areas.append("南区")
        if any(str(postal_code).startswith(str(code)[:2]) for code in west_area):
            possible_areas.append("西区")
        if any(str(postal_code).startswith(str(code)[:2]) for code in city_center):
            possible_areas.append("市中心")

        if possible_areas:
            return f"未知区域，可能在以下区域之一: {', '.join(possible_areas)}"
        else:
            return "未知区域"

# 从 Excel 文件读取邮编数据
def process_excel(input_file, output_file):
    df = pd.read_excel(input_file)

    # 确保邮编列存在
    if '收件人编码' not in df.columns:
        raise ValueError("Excel 文件中没有名为 '收件人编码' 的列")

    # 处理每个邮编并添加区域列
    df['区域'] = df['收件人编码'].apply(find_area)

    # 将结果写入新的 Excel 文件
    df.to_excel(output_file, index=False)

# 示例使用
input_file = '第二十七批次.xlsx'  # 输入的 Excel 文件
output_file = '邮编对应区域.xlsx'  # 输出的 Excel 文件

process_excel(input_file, output_file)