import pandas as pd
import openai
import json
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import os
import requests

# ========== 设置 OpenAI API Key ==========
client = openai.OpenAI(api_key="ChangeKey")  # 替换为你自己的 key

# ========== 读取 Excel 文件 ==========
excel_input_path = "实习面试题.xlsx"  # 确保文件和脚本同目录
Sumfile = pd.read_excel(excel_input_path, header=0, sheet_name="题目内容")
Sumfile.columns = Sumfile.columns.str.strip().str.replace(" ", "_").str.replace("-", "_")
Sumfile = Sumfile.fillna("WU")
data_list = Sumfile.to_dict(orient='records')

# ========== 描述字段 → 图像字段 ==========
image_fields = {
    "QUESTION_PIC_DESCRIPTION": "QUESTION_PIC",
    "PIC_DESCRIPTION_for_OPTION_A": "PIC_for_OPTION_A",
    "PIC_DESCRIPTION_for_OPTION_B": "PIC_for_OPTION_B",
    "PIC_DESCRIPTION_for_OPTION_C": "PIC_for_OPTION_C",
    "PIC_DESCRIPTION_for_OPTION_D": "PIC_for_OPTION_D"
}

# ========== 生成模拟题 + 图片 ==========
new_data_list = []

for i, base_question in enumerate(data_list[:5]):
    print(f"\n📌 正在为第 {i+1} 题生成模拟题...")
    for j in range(3):
        try:
            # ✅ 使用安全方式构建 prompt，防止编码错误
            prompt = "\n".join([
                "请根据以下题目信息，生成一个风格一致的“全新模拟题”。结构、字段格式保持一致，题目内容需不同。返回 JSON 格式的题目信息字典。",
                "原始题目信息如下：",
                json.dumps(base_question, ensure_ascii=False),
                "只返回新的 JSON 字典，不要解释。"
            ])

            # GPT 调用
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )

            new_question = json.loads(response.choices[0].message.content)

            # 生成图像并插入路径
            for desc_field, image_field in image_fields.items():
                description = new_question.get(desc_field, None)
                if description and description != "WU":
                    image_gen_response = client.images.generate(
                        model="dall-e-3",
                        prompt=description,
                        n=1,
                        size="1024x1024"
                    )
                    image_url = image_gen_response.data[0].url

                    # ✅ 使用 client 下载图像
                    image_data = requests.get(image_url).content 

                    # 保存本地图片
                    image_name = f"{i+1}_{j+1}_{image_field}.png"
                    with open(image_name, "wb") as f:
                        f.write(image_data)

                    new_question[image_field] = image_name
                else:
                    new_question[image_field] = "WU"

            new_data_list.append(new_question)
            print(f"✅ 第 {i+1} 题 - 模拟题 {j+1} 生成成功")

        except Exception as e:
            print(f"❌ 生成失败：{e}")

# ========== 写入 Excel 并插图 ==========
if not new_data_list:
    print("❌ 未生成任何题目，请检查 API 调用是否正常")
    exit()

wb = Workbook()
ws = wb.active
ws.title = "题目内容"

headers = list(new_data_list[0].keys())
ws.append(headers)

for row_idx, item in enumerate(new_data_list, start=2):
    row_data = []
    for col_idx, header in enumerate(headers, start=1):
        value = item[header]
        if isinstance(value, str) and value.endswith(".png") and os.path.exists(value):
            img = XLImage(value)
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 25
            ws.row_dimensions[row_idx].height = 90
            ws.add_image(img, f"{col_letter}{row_idx}")
            row_data.append("")
        else:
            row_data.append(value)
    ws.append(row_data)

# 保存结果
output_path = "生成的新题目合集_带图.xlsx"
wb.save(output_path)
print(f"\n✅ 所有题目和图片已生成，文件保存为：{output_path}")
