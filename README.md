# 浙江省电梯招投标信息自动采集工具

自动从浙江省公共资源交易平台 (ggzy.zj.gov.cn) 采集电梯相关招投标信息，输出带颜色分类的 Excel 报告。

## 功能特性

- **6组关键词自动搜索**：电梯、电梯采购、电梯维保、电梯维修、电梯改造、加装电梯
- **智能分类**：自动将标讯分为新梯采购 / 维保服务 / 维修改造 / 旧梯加装
- **预算提取**：从标题和内容中正则匹配预算金额
- **Excel彩色报表**：不同业务类型用不同颜色区分，一目了然
- **完全免费**：调用的是公共平台API，零成本运行

## 输出示例

| 序号 | 发布时间 | 地区 | 项目名称 | 业务类型 | 公告类别 | 详情链接 | 预算(万元) |
|------|---------|------|---------|---------|---------|---------|-----------|
| 1 | 2026-06-14 | 杭州市 | XX小区电梯设备采购项目 | 新梯采购 | 招标公告 | [链接] | 350 |
| 2 | 2026-06-14 | 宁波市 | XX大厦电梯年度维保服务 | 维保服务 | 招标公告 | [链接] | 28 |

Excel 中不同业务类型有不同底色：
- 🔵 新梯采购 — 浅蓝色
- 🟢 维保服务 — 浅绿色
- 🟠 维修改造 — 浅橙色
- 🔴 旧梯加装 — 浅红色

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/lisazhang0707-ops/zj-elevator-bid-collector.git
cd zj-elevator-bid-collector
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行采集

```bash
# 默认采集最近7天
python collector.py

# 采集最近3天
python collector.py --days 3

# 指定输出目录
python collector.py --output ./my_data

# 自定义关键词
python collector.py --keywords "电梯,自动扶梯,观光电梯"
```

### 4. 查看结果

Excel 报告保存在 `./output/` 目录下，文件名格式：`浙江电梯标讯_YYYY-MM-DD.xlsx`

## 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--days` | 采集最近几天的标讯 | 7 |
| `--output` | 输出目录路径 | ./output |
| `--keywords` | 搜索关键词（逗号分隔） | 电梯,电梯采购,电梯维保,电梯维修,电梯改造,加装电梯 |

## 定时自动运行

### Windows 任务计划程序

```cmd
# 每天早上8点自动运行
schtasks /create /tn "电梯标讯采集" /tr "python C:\path\to\collector.py" /sc daily /st 08:00
```

### macOS / Linux Crontab

```bash
# 每天早上8点运行
0 8 * * * cd /path/to/zj-elevator-bid-collector && python collector.py >> collect.log 2>&1
```

## 数据源

浙江省公共资源交易平台：https://ggzy.zj.gov.cn/jyxxgk/list.html

## 技术说明

- 使用平台内部 REST API (`getFullTextDataNew`) 获取结构化数据
- 自动分页获取，每组关键词最多获取3页（150条）
- 内置去重逻辑（基于项目ID + 标题），避免重复采集
- 自动排除不相关项目（监理、设计、咨询、检测等）

## 依赖

- Python 3.8+
- requests
- pandas
- openpyxl

## License

MIT
