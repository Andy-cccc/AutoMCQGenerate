# AutoMCQGenerate
题模拟生成器（Excel + OpenAI）

本项目用于从 Excel 中读取原始题，通过调用 OpenAI API（GPT-3.5）自动生成风格一致的模拟题，并将所有题目保存为新的 Excel 文件，支持图片与字段扩展。

功能亮点:
从 Excel 表格读取题目数据
自动调用 GPT 接口生成多个风格一致的新题
保留原始结构，兼容图片字段
支持保存生成结果为 Excel 文件
可拓展为完整的题库系统

使用的技术栈:
pandas 读取/处理 Excel 数据  
openai 调用 ChatGPT 接口生成新题 
json      结构化处理题目信息

 使用说明:
1. 安装依赖库
bash
pip install pandas openai openpyxl
2.准备 Excel 文件(已提供例题)
命名为 实习面试题.xlsx
格式为(列名)
GRADE	LEVEL	Knowledge Point	MULTIPLE-CHOICE QUESTION	ANSWER	QUESTION PIC DESCRIPTION	 PIC DESCRIPTION for OPTION A 	 PIC DESCRIPTION for OPTION B 	 PIC DESCRIPTION for OPTION C	 PIC DESCRIPTION for OPTION D	QUESTION PIC	 PIC for OPTION A 	 PIC for OPTION B 	 PIC for OPTION C	 PIC for OPTION D
3.运行Final.py

输出结果:
成功运行后，会生成：
新的 Excel 文件：生成的新题目合集.xlsx
原始题 + 模拟题汇总，共 20 题（前 5 题位置不变，第一题拓展的三题在第五题后面三行，以此类推，新题15题，原题5题）
