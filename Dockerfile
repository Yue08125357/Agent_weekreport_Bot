FROM python:3.11-slim

WORKDIR /app

# 安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 复制代码和模板
COPY web_app.py .
COPY template.docx .

# 暴露端口
EXPOSE 8000

# 启动命令
CMD ["uvicorn", "web_app:app", "--host", "0.0.0.0", "--port", "8000"]
