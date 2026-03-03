"""
周报助手 - 云端版
用户复制表格内容，网页自动生成周报
"""

import io
import sys
import csv
import re
from pathlib import Path
from datetime import datetime, timedelta
from io import StringIO

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from fastapi import FastAPI
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
            max-width: 800px;
            width: 100%;
        }
        h1 { color: #333; margin-bottom: 10px; font-size: 28px; text-align: center; }
        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; text-align: center; }
        
        .step {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            border-left: 4px solid #667eea;
        }
        .step-title {
            font-weight: 600;
            color: #333;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .step-num {
            background: #667eea;
            color: white;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
        }
        
        .link-box {
            background: #e3f2fd;
            padding: 12px 16px;
            border-radius: 8px;
            margin: 10px 0;
            word-break: break-all;
        }
        .link-box a {
            color: #1565c0;
            text-decoration: none;
        }
        .link-box a:hover {
            text-decoration: underline;
        }
        
        label { display: block; font-weight: 600; color: #333; margin-bottom: 8px; margin-top: 15px; }
        
        textarea {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 14px;
            resize: vertical;
            min-height: 120px;
            font-family: inherit;
        }
        textarea:focus { outline: none; border-color: #667eea; }
        
        input[type="text"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
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
            margin-top: 20px;
        }
        .btn:hover { transform: translateY(-2px); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        
        .msg {
            padding: 15px;
            border-radius: 8px;
            margin-top: 20px;
            text-align: center;
        }
        .msg.info { background: #e3f2fd; color: #1565c0; }
        .msg.success { background: #e8f5e9; color: #2e7d32; }
        .msg.error { background: #ffebee; color: #c62828; }
        
        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 10px auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        
        .tip {
            color: #888;
            font-size: 12px;
            margin-top: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>周报助手</h1>
        <p class="subtitle">研究管理团队周报自动生成</p>
        
        <div id="msg" class="msg info" style="display:none;"></div>
        
        <!-- 步骤1 -->
        <div class="step">
            <div class="step-title">
                <span class="step-num">1</span>
                打开WPS表格
            </div>
            <div class="link-box">
                <a href="WPS_URL" target="_blank" id="wps_link">WPS_URL</a>
            </div>
            <p style="color:#666;font-size:13px;">点击上方链接，如需登录请扫码</p>
        </div>
        
        <!-- 步骤2 -->
        <div class="step">
            <div class="step-title">
                <span class="step-num">2</span>
                复制表格内容
            </div>
            <p style="color:#666;font-size:13px;margin-bottom:10px;">
                在WPS中：选择最新Sheet → Ctrl+A 全选 → Ctrl+C 复制
            </p>
        </div>
        
        <!-- 步骤3 -->
        <div class="step">
            <div class="step-title">
                <span class="step-num">3</span>
                粘贴并生成
            </div>
            
            <label>粘贴表格内容</label>
            <textarea id="content" placeholder="在此粘贴从WPS复制的内容（Ctrl+V）..."></textarea>
            <p class="tip">直接 Ctrl+V 粘贴即可</p>
            
            <label>日期范围</label>
            <input type="text" id="date" placeholder="如：0303-0307（留空自动识别）">
            
            <button type="button" class="btn" id="btn" onclick="generate()">生成周报</button>
        </div>
        
        <div id="loading" style="display:none;text-align:center;padding:20px;">
            <div class="spinner"></div>
            <p>正在生成周报...</p>
        </div>
    </div>
    
    <script>
        function generate() {
            const content = document.getElementById('content').value.trim();
            const date = document.getElementById('date').value.trim();
            const msgDiv = document.getElementById('msg');
            const loadingDiv = document.getElementById('loading');
            const btn = document.getElementById('btn');
            
            if (!content) {
                msgDiv.className = 'msg error';
                msgDiv.textContent = '请先粘贴表格内容';
                msgDiv.style.display = 'block';
                return;
            }
            
            btn.disabled = true;
            loadingDiv.style.display = 'block';
            msgDiv.style.display = 'none';
            
            fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ content: content, date: date })
            })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(data => { throw new Error(data.error || '生成失败'); });
                }
                return response.blob();
            })
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = '周报(' + (date || '未知日期') + ').docx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                msgDiv.className = 'msg success';
                msgDiv.textContent = '周报已生成并下载！';
                msgDiv.style.display = 'block';
            })
            .catch(err => {
                msgDiv.className = 'msg error';
                msgDiv.textContent = '错误: ' + err.message;
                msgDiv.style.display = 'block';
            })
            .finally(() => {
                btn.disabled = false;
                loadingDiv.style.display = 'none';
            });
        }
    </script>
</body>
</html>
""".replace("WPS_URL", DEFAULT_WPS_URL)


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


class GenerateRequest(BaseModel):
    content: str
    date: str = ""


@app.post("/api/generate")
async def generate(req: GenerateRequest):
    """生成周报"""
    try:
        if not req.content:
            return JSONResponse({"error": "表格内容为空"}, status_code=400)
        
        # 如果没有日期，尝试从内容中提取
        date = req.date
        if not date:
            # 尝试匹配日期格式
            match = re.search(r'(\d{4}[-—]\d{4})', req.content)
            if match:
                date = match.group(1)
            else:
                date = datetime.now().strftime("%m%d") + "-" + (datetime.now() + timedelta(days=4)).strftime("%m%d")
        
        doc_bytes = generate_report(req.content, date)
        
        return Response(
            content=doc_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="周报({date}).docx"'}
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
