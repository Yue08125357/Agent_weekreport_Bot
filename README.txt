# 周报自动汇总助手

## 快速使用

### 方法一：双击启动
直接双击 `启动周报助手.bat` 即可！

### 方法二：命令行启动
```bash
cd d:/MyProject/Agent_weekreport_Bot
python weekreport_bot.py
```

## 使用流程

1. **启动程序** → 自动打开浏览器访问WPS表格
2. **扫码登录** → 用微信扫码登录（只需一次）
3. **自动读取** → 程序自动读取最新Sheet数据
4. **生成周报** → 输出到 `output` 文件夹

## 文件说明

```
Agent_weekreport_Bot/
├── 启动周报助手.bat    ← 双击这个启动
├── weekreport_bot.py   ← 主程序
├── output/             ← 生成的周报存放位置
└── README.txt

demo_week/              ← 历史周报模板库
├── 研究管理团队周报（0104-0109).docx
└── 研究管理团队周报（0126-0130).docx
```

## 常见问题

Q: 首次运行很慢？
A: 首次需要下载浏览器，约100MB，之后秒开

Q: 表格读取不完整？
A: 程序会自动截图保存，可手动补充

Q: 如何修改表格链接？
A: 编辑 weekreport_bot.py 第25行 WPS_TABLE_URL
