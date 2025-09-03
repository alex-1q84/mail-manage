# 项目架构文档

## 项目概述
本项目是一个邮件管理工具，主要用于检查邮件附件大小并管理邮件数据。

## 目录结构
```
/Users/mywo/lab/mail-manage
├── .env                # 环境变量配置文件
├── .envrc              # 环境配置脚本
├── .gitignore          # Git 忽略规则
├── .python-version     # Python 版本配置
├── pyproject.toml      # Python 项目配置
├── check_large_attachments.py  # 检查邮件附件的脚本
├── mails.txt           # 邮件数据文件
├── README.md           # 项目说明（当前为空）
└── uv.lock             # 依赖锁定文件
```

## 功能模块
1. **check_large_attachments.py**
   - 功能：检查邮件附件大小，确保符合预设的限制。
   - 依赖：Python 环境，配置文件 `.env` 中的参数。

2. **mails.txt**
   - 功能：存储邮件数据，供脚本处理。

## 配置说明
- `.env`: 包含环境变量，如附件大小限制等。
- `pyproject.toml`: 定义项目依赖和构建配置。

## 后续建议
1. 完善 `README.md`，补充项目使用说明和开发指南。
2. 扩展 `check_large_attachments.py` 功能，支持更多邮件管理操作。
3. 考虑将邮件数据存储到数据库（如 SQLite）以提高管理效率。