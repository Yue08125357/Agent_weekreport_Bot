"""
周报助手 - 云端全自动版
用户只需扫码登录，其他全自动
"""

import asyncio
import io
import sys
import csv
import re
import os
from pathlib import Path
from datetime import datetime, timedelta
from io import StringIO
from typing import Optional

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, Response, JSONResponse
from pydantic import BaseModel

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 配置
TEMPLATE = Path("template.docx")
DEFAULT_WPS_URL = "https://www.kdocs.cn/l/cmvjNWclJM5P"

app = FastAPI(title="周报助手")

# 任务存储
tasks = {}


def parse_table(text):
    """解析表格"""
    reader = csv.reader(StringIO(text), delimiter='\t')
    rows = []
    for row in reader:
        rows.append([cell.strip() for cell in row])
    return rows


def set_font(cell, font='仿宋', size=11, bold=False):
    """设置字体"""
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.name = font
            r._element.rPr.rFonts.set(qn('w:eastAsia'), font)
            r.font.size = Pt(size)
            r.font.bold = bold


def add_borders(table):
    """加边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr or OxmlElement('w:tblPr')
    borders = OxmlElement('w:tblBorders')
    for n in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{n}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        borders.append(b)
    tblPr.append(borders)
    if not tbl.tblPr:
        tbl.insert(0, tblPr)


def generate_report(table_content: str, date: str) -> bytes:
    """生成周报文档"""
    data = parse_table(table_content)
    if not data:
        raise ValueError("表格内容为空")
    
    doc = Document(str(TEMPLATE))
    
    # 删除旧表格
    for t in list(doc.tables):
        t._tbl.getparent().remove(t._tbl)
    
    # 修改标题日期
    for para in doc.paragraphs:
        full_text = ''.join([r.text for r in para.runs])
        if '工作周报' in full_text:
            for run in para.runs:
                run.text = ''
            para.runs[0].text = f'工作周报（{date}）'
            para.runs[0].font.name = '微软雅黑'
            para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
            para.runs[0].font.size = Pt(20)
            para.runs[0].font.bold = True
            break
    
    # 找详细工单位置
    detail = None
    for para in doc.paragraphs:
        if '详细工单' in para.text:
            detail = para
            break
    
    cols = max(len(r) for r in data)
    
    # 表格1
    t1 = doc.add_table(len(data), cols)
    add_borders(t1)
    for i, row in enumerate(data):
        for j, txt in enumerate(row):
            t1.rows[i].cells[j].text = txt
            set_font(t1.rows[i].cells[j], bold=(i==0))
    
    if detail:
        detail._element.addprevious(t1._tbl)
    
    # 表格2
    data2 = data + [['其他需汇报信息：无', '', '', '']]
    t2 = doc.add_table(len(data2), cols)
    add_borders(t2)
    for i, row in enumerate(data2):
        for j, txt in enumerate(row):
            t2.rows[i].cells[j].text = txt
            set_font(t2.rows[i].cells[j], bold=(i==0))
    
    if detail:
        detail._element.addnext(t2._tbl)
    
    # 保存到内存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


async def run_browser_task(task_id: str, wps_url: str, date: str):
    """后台运行浏览器任务"""
    from playwright.async_api import async_playwright, TimeoutError
    
    try:
        tasks[task_id]["status"] = "launching"
        tasks[task_id]["message"] = "正在启动浏览器..."
        
        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        
        # 打开WPS
        tasks[task_id]["status"] = "loading"
        tasks[task_id]["message"] = "正在打开WPS表格..."
        await page.goto(wps_url, wait_until="networkidle")
        
        # 等待一下让页面加载
        await asyncio.sleep(3)
        
        # 截图显示当前页面（登录二维码）
        tasks[task_id]["status"] = "waiting_login"
        tasks[task_id]["message"] = "请扫码登录微信"
        
        # 持续截图，等待登录
        for i in range(60):  # 最多等2分钟
            screenshot = await page.screenshot(type="jpeg", quality=50)
            screenshot_base64 = screenshot.hex()
            tasks[task_id]["screenshot"] = screenshot.hex()
            tasks[task_id]["wait_time"] = i * 2
            
            # 检查是否登录成功（URL变化或检测到表格元素）
            current_url = page.url
            if 'passport' not in current_url and 'singlesign' not in current_url:
                # 检查是否有表格元素
                try:
                    table_elements = await page.query_selector_all('[class*="sheet"], [class*="editor"], [class*="canvas"]')
                    if table_elements:
                        tasks[task_id]["message"] = "登录成功，正在读取表格..."
                        break
                except:
                    pass
            
            await asyncio.sleep(2)
        
        # 登录后继续
        tasks[task_id]["status"] = "reading"
        tasks[task_id]["message"] = "正在读取表格数据..."
        
        # 等待表格完全加载
        await asyncio.sleep(3)
        
        # 获取最新Sheet名称
        sheet_name = ""
        for selector in ['[class*="sheet-tab"]', '.sheet-tab', '[role="tab"]']:
            try:
                tabs = await page.query_selector_all(selector)
                if tabs:
                    last_tab = tabs[-1]
                    sheet_name = await last_tab.inner_text()
                    sheet_name = sheet_name.strip()
                    await last_tab.click()
                    await asyncio.sleep(2)
                    break
            except:
                continue
        
        # 全选复制
        tasks[task_id]["message"] = "正在复制表格..."
        try:
            await page.click('[class*="canvas"], [class*="editor"], [class*="sheet"]')
        except:
            pass
        await asyncio.sleep(0.5)
        await page.keyboard.press('Control+A')
        await asyncio.sleep(0.5)
        await page.keyboard.press('Control+C')
        await asyncio.sleep(1)
        
        # 使用JavaScript读取选中内容
        table_text = await page.evaluate('''() => {
            const selection = window.getSelection();
            if (selection.rangeCount > 0) {
                const range = selection.getRangeAt(0);
                const div = document.createElement('div');
                div.appendChild(range.cloneContents());
                
                // 转换为表格格式
                const cells = div.querySelectorAll('td, th, [class*="cell"]');
                if (cells.length > 0) {
                    let result = [];
                    let currentRow = [];
                    let lastY = -1;
                    
                    cells.forEach(cell => {
                        const rect = cell.getBoundingClientRect();
                        const text = cell.innerText?.trim() || '';
                        
                        if (lastY !== -1 && Math.abs(rect.y - lastY) > 15) {
                            if (currentRow.length > 0) result.push(currentRow);
                            currentRow = [];
                        }
                        
                        currentRow.push(text);
                        lastY = rect.y;
                    });
                    
                    if (currentRow.length > 0) result.push(currentRow);
                    return result.map(row => row.join('\\t')).join('\\n');
                }
                
                return div.innerText;
            }
            return '';
        }''')
        
        # 关闭浏览器
        await browser.close()
        await playwright.stop()
        
        # 确定日期
        if not date:
            match = re.search(r'(\d{4}[-—]\d{4})', sheet_name)
            date = match.group(1) if match else sheet_name
        
        tasks[task_id]["date"] = date
        tasks[task_id]["message"] = "正在生成周报..."
        
        # 生成报告
        if not table_text:
            raise ValueError("未能读取表格内容")
        
        doc_bytes = generate_report(table_text, date)
        tasks[task_id]["document"] = doc_bytes.hex()
        tasks[task_id]["status"] = "done"
        tasks[task_id]["message"] = "周报已生成！"
        
    except Exception as e:
        tasks[task_id]["status"] = "error"
        tasks[task_id]["message"] = f"错误: {str(e)}"


# HTML页面
HTML_PAGE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>周报助手</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 700px;
            width: 100%;
        }
        h1 { color: #333; margin-bottom: 10px; font-size: 28px; text-align: center; }
        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; text-align: center; }
        label { display: block; font-weight: 600; color: #333; margin-bottom: 8px; }
        input[type="text"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            margin-bottom: 20px;
        }
        input[type="text"]:focus { outline: none; border-color: #667eea; }
        .btn {
            width: 100%;
            padding: 14px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s;
        }
        .btn:hover { transform: translateY(-2px); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        .msg {
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            text-align: center;
        }
        .msg.info { background: #e3f2fd; color: #1565c0; }
        .msg.success { background: #e8f5e9; color: #2e7d32; }
        .msg.error { background: #ffebee; color: #c62828; }
        .screenshot-container {
            margin: 20px 0;
            text-align: center;
            display: none;
        }
        .screenshot-container img {
            max-width: 100%;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        }
        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .progress { display: none; text-align: center; padding: 20px; }
 }
    </style>
</head>
<body>
    <div class="container">
        <h1>周报助手</h1>
        <p class="subtitle">研究管理团队周报自动生成</p>
        
        <div id="msg" class="msg info" style="display:none;"></div>
        
        <form id="form">
            <label>WPS表格地址</label>
            <input type="text" id="wps_url" value="WPS_URL" placeholder="WPS在线表格地址">
            
            <label>日期范围（如：0303-0307）</label>
            <input type="text" id="date" placeholder="留空自动从Sheet名称识别">
            
            <button type="submit" class="btn" id="btn">生成周报</button>
        </form>
        
        <div class="progress" id="progress">
            <div class="spinner"></div>
            <div id="status">正在处理...</div>
        </div>
        
        <div class="screenshot-container" id="screenshot-container">
            <p style="margin-bottom:10px;color:#666;">请扫描下方二维码登录微信</p>
            <img id="screenshot" src="" alt="登录二维码">
        </div>
    </div>
    
    <script>
        document.getElementById('form').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const wpsUrl = document.getElementById('wps_url').value.trim();
            const date = document.getElementById('date').value.trim();
            const msgDiv = document.getElementById('msg');
            const progressDiv = document.getElementById('progress');
            const statusDiv = document.getElementById('status');
            const screenshotContainer = document.getElementById('screenshot-container');
            const screenshotImg = document.getElementById('screenshot');
            const btn = document.getElementById('btn');
            
            btn.disabled = true;
            msgDiv.style.display = 'none';
            progressDiv.style.display = 'block';
            screenshotContainer.style.display = 'none';
            
            try {
                // 开始生成
                statusDiv.textContent = '正在启动...';
                const startRes = await fetch('/api/start', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ wps_url: wpsUrl, date: date })
                });
                const startData = await startRes.json();
                const taskId = startData.task_id;
                
                // 轮询状态
                let done = false;
                while (!done) {
                    await new Promise(r => setTimeout(r, 1500));
                    
                    const checkRes = await fetch('/api/status/' + taskId);
                    const data = await checkRes.json();
                    
                    statusDiv.textContent = data.message || '处理中...';
                    
                    if (data.status === 'waiting_login') {
                        progressDiv.style.display = 'none';
                        screenshotContainer.style.display = 'block';
                        if (data.screenshot) {
                            screenshotImg.src = 'data:image/jpeg;base64,' + hexToBase64(data.screenshot);
                        }
                    } else if (data.status === 'done') {
                        done = true;
                        screenshotContainer.style.display = 'none';
                        progressDiv.style.display = 'none';
                        
                        // 下载文档
                        window.location.href = '/api/download/' + taskId;
                        
                        msgDiv.className = 'msg success';
                        msgDiv.textContent = '周报已生成，正在下载...';
                        msgDiv.style.display = 'block';
                        
                    } else if (data.status === 'error') {
                        throw new Error(data.message);
                    }
                }
                
            } catch (err) {
                msgDiv.className = 'msg error';
                msgDiv.textContent = '错误: ' + err.message;
                msgDiv.style.display = 'block';
                progressDiv.style.display = 'none';
                screenshotContainer.style.display = 'none';
            } finally {
                btn.disabled = false;
            }
        });
        
        function hexToBase64(hex) {
            const bytes = new Uint8Array(hex.match(/.{2}/g).map(byte => parseInt(byte, 16)));
            let binary = '';
            bytes.forEach(b => binary += String.fromCharCode(b));
            return btoa(binary);
        }
    </script>
</body>
</html>
""".replace("WPS_URL", DEFAULT_WPS_URL)


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


class StartRequest(BaseModel):
    wps_url: str = DEFAULT_WPS_URL
    date: str = ""


@app.post("/api/start")
async def start_task(req: StartRequest):
    """开始任务"""
    task_id = datetime.now().strftime("%Y%m%d%H%M%S") + str(os.getpid())
    tasks[task_id] = {
        "status": "starting",
        "message": "正在初始化..."
    }
    
    # 启动后台任务
    asyncio.create_task(run_browser_task(task_id, req.wps_url, req.date))
    
    return {"task_id": task_id}


@app.get("/api/status/{task_id}")
async def get_status(task_id: str):
    """获取任务状态"""
    task = tasks.get(task_id, {"status": "unknown", "message": "任务不存在"})
    # 不返回 document 字段（太大）
    result = {k: v for k, v in task.items() if k != "document"}
    return result


@app.get("/api/download/{task_id}")
async def download(task_id: str):
    """下载文档"""
    task = tasks.get(task_id, {})
    if task.get("document"):
        doc_hex = task["document"]
        doc_bytes = bytes.fromhex(doc_hex)
        date = task.get("date", "unknown")
        
        return Response(
            content=doc_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="周报({date}).docx"'}
        )
    return {"error": "文档不存在"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
