FROM python:3.8-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 创建非root用户
RUN useradd -m -u 1000 appuser

# 创建必要的目录并设置权限
RUN mkdir -p /app/output && chown -R appuser:appuser /app

# 只复制必要的文件
COPY --chown=appuser:appuser requirements-prod.txt .
COPY --chown=appuser:appuser src/ src/
COPY --chown=appuser:appuser config/ config/
COPY --chown=appuser:appuser main.py .
COPY --chown=appuser:appuser api_app.py .

# 安装Python依赖（只安装生产环境依赖）
RUN pip install --no-cache-dir -r requirements-prod.txt

# 切换到非root用户
USER appuser

# 暴露端口
EXPOSE 8000

# 设置环境变量
ENV PYTHONUNBUFFERED=1

# 创建数据卷
VOLUME ["/app/output"]

# 运行FastAPI应用
CMD ["uvicorn", "api_app:app", "--host", "0.0.0.0", "--port", "8000"]