from fastapi import FastAPI, Body
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import pandas as pd
import uuid
import os

# 这里导入你的脚本，如果名字不对你自己改一下
try:
    from generate_ppt import generate_ppt
except ImportError:
    # 如果没找到你的脚本，临时弄个空的防止报错退出
    def generate_ppt(excel_path, ppt_path):
        pass

app = FastAPI()

# 允许前端跨域访问（本地调试方便）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 创建临时文件夹
os.makedirs("temp_files", exist_ok=True)

# 前端静态文件（index.html 在项目根目录）
app.mount("/static", StaticFiles(directory="."), name="static")

# 首页：直接返回 index.html
@app.get("/")
async def index():
    if not os.path.exists("index.html"):
        return {"status": "error", "message": "找不到 index.html"}
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

# 【新增这几行】：给前端提供固定模板
@app.get("/api/template")
async def get_excel_template():
    # 确保你的 template.xlsx 和 main.py 在同一个文件夹里
    if not os.path.exists("template.xlsx"):
        return {"status": "error", "message": "找不到模板文件 template.xlsx"}
    return FileResponse(path="template.xlsx", filename="template.xlsx")

# 核心修改点：在括号里加上 data: list = Body(...)，这样网页上就会出现输入框了！
@app.post("/api/generate-ppt")
async def api_generate_ppt(
    data: list = Body(..., example=[{"姓名": "张三", "业绩": 100}, {"姓名": "李四", "业绩": 200}])
):

    task_id = str(uuid.uuid4())
    temp_excel_path = f"temp_files/{task_id}.xlsx"
    output_ppt_path = f"temp_files/{task_id}.pptx"

    # 把传来的数据存成 Excel
    df = pd.DataFrame(data)
    df.to_excel(temp_excel_path, index=False)

    # 调用你的 PPT 生成脚本
    try:
        generate_ppt(temp_excel_path, output_ppt_path)
    except Exception as e:
        return {"status": "error", "message": f"生成失败，报错信息: {str(e)}"}

    # 如果你的脚本还没连上，没关系，我们先把刚才生成的 Excel 弹出来给你下载，证明接口通了！
    if not os.path.exists(output_ppt_path):
        return FileResponse(temp_excel_path, filename="测试数据转成的Excel.xlsx")

    # 如果有 PPT，就返回 PPT 下载
    return FileResponse(
        path=output_ppt_path,
        filename="最终生成的报告.pptx"
    )
