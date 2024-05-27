import os
from typing import Optional
from xml.dom.minidom import Document
from fastapi import FastAPI, Form, HTTPException, Path, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import pandas as pd


app = FastAPI()


class CompanyRequest(BaseModel):
    company_name: str


class Item(BaseModel):
    file: Optional[UploadFile] = None
    key1: Optional[str] = None
    val1: Optional[str] = None


def calculate(operation_str, number):
    # 确定运算符和操作数
    operator = operation_str[0]  # 运算符是第一个字符
    operand = float(
        operation_str[1:]
    )  # 操作数是从第二个字符开始的部分，转换为浮点数以处理所有情况

    # 根据运算符进行相应的计算
    if operator == "+":
        result = number + operand
    elif operator == "-":
        result = number - operand
    elif operator == "*":
        result = number * operand
    elif operator == "/":
        if operand == 0:
            raise ValueError("Cannot divide by zero")
        result = number / operand
    else:
        raise ValueError("Invalid operation")

    return result


# 將 /image 資料夾掛載到 /static 路徑
app.mount("/image", StaticFiles(directory="image"), name="image")


# 首頁設定成某個html
@app.get("/", response_class=FileResponse)
async def read_root():
    return FileResponse("./public/index.html")


@app.post("/get_company_data")
def get_company_data(request: CompanyRequest):
    print("********")


@app.post("/test")
async def test333(
    file: Optional[UploadFile] = Form(None),
    key1: Optional[str] = Form(None),
    val1: Optional[str] = Form(None),
):
    item = Item(
        file=file,
        key1=key1,
        val1=val1,
    )
    print(item)

    extension = os.path.splitext(item.file.filename)[1].lower()
    file_n = os.path.splitext(item.file.filename)[0].lower()

    if extension not in [".xlsx"]:
        raise HTTPException(
            status_code=400,
            detail="Invalid file extension. Only .xlsx files are allowed.",
        )

    file_path = "image/" + file_n + "_afterFix" + ".xlsx"  # 文件保存位置
    with open(file_path, "wb") as f:
        f.write(await file.read())  # 寫入文件

    print(file_path)

    # 讀取 Excel 文件
    df = pd.read_excel(file_path)

    # 查找 "薪水" 列並將其數值乘以五
    if item.key1 in df.columns:
        df[item.key1] = calculate(item.val1, df[item.key1])

    # 保存修改後的數據到新文件
    new_file_path = "image/" + file_n + "_afterFix" + ".xlsx"
    df.to_excel(new_file_path, index=False)

    word_path = "image/" + file_n + "_afterFix" + ".xlsx"
    return {"screenshot_path": f"{word_path}"}
