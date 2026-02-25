# 债券买家匹配 (Bond Buyer Match)

基于 **iFinD（同花顺）** 数据源的债券信息查询与买家匹配工具。

## 使用前准备

1. **安装 iFinD 接口**  
   - 双击运行 `setup_ifind.bat`，或命令行执行：`pip install iFinDAPI`  
   - 文档与账号申请：[同花顺数据接口](https://quantapi.51ifind.com/)

2. **配置账号**  
   - 启动程序后，打开 **帮助 → iFinD 设置**  
   - 填写同花顺数据接口的**账号**和**密码**并确定  

   命令行查询时，也可设置环境变量：`IFIND_USER`、`IFIND_PASSWORD`。

## 依赖

- Python 3.8+
- PyQt5、pandas、openpyxl（见 `requirements.txt`）
- **iFinD**：`pip install iFinDAPI`，并在「iFinD 设置」中填写账号密码

## 项目结构

```
bond-match/
├── app.py              # 主程序（PyQt5 界面）
├── ifind_api.py        # iFinD 数据接口
├── data_provider.py    # 数据 API 入口（iFinD）
├── bond_lookup.py      # 命令行债券查询
├── import_data.py      # Excel 数据导入
├── database/
│   ├── db_manager.py   # 数据库管理
│   └── schema.sql      # 表结构
├── data/               # 放置 holdings / can_buy_lists / contacts 的 Excel
├── setup_ifind.bat     # 安装 iFinD 接口
├── requirements.txt
└── README.md
```
