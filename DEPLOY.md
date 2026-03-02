# 部署指南

## 方法一：Railway 部署（推荐，免费）

### 步骤1: 注册 GitHub 账号
如果还没有 GitHub 账号，先注册一个：https://github.com

### 步骤2: 上传代码到 GitHub
1. 在 GitHub 创建新仓库（如 `weekreport-bot`）
2. 上传以下文件：
   - `web_app.py`
   - `requirements.txt`
   - `Dockerfile`
   - `template.docx`
   - `railway.json`

### 步骤3: 部署到 Railway
1. 访问 https://railway.app
2. 点击 "Start a New Project"
3. 选择 "Deploy from GitHub repo"
4. 授权 Railway 访问你的 GitHub
5. 选择 `weekreport-bot` 仓库
6. 点击 "Deploy Now"

### 步骤4: 获取访问地址
部署完成后，Railway 会提供一个访问地址，如：
`https://weekreport-bot.up.railway.app`

---

## 方法二：Render 部署（免费但有冷启动）

### 步骤1-2: 同上，上传到 GitHub

### 步骤3: 部署到 Render
1. 访问 https://render.com
2. 点击 "New" → "Web Service"
3. 连接 GitHub 仓库
4. 选择 `weekreport-bot` 仓库
5. 配置：
   - Environment: Docker
   - 点击 "Create Web Service"

### 步骤4: 获取访问地址
部署完成后，地址如：
`https://weekreport-bot.onrender.com`

---

## 方法三：内网穿透（最快，需保持电脑开机）

### 使用 ngrok
```bash
# 安装 ngrok
pip install pyngrok

# 启动本地服务
python web_app.py

# 新开终端，运行 ngrok
ngrok http 8000
```

会得到一个公网地址，如 `https://xxxx.ngrok.io`

---

## 文件清单

需要上传的文件：
- web_app.py (主程序)
- requirements.txt (依赖)
- Dockerfile (容器配置)
- template.docx (模板文件)
- railway.json (Railway配置)
