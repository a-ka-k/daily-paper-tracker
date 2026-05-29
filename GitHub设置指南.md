# GitHub Actions 部署指南

## 完整部署步骤

### 第一步：创建 GitHub 仓库

1. 登录 [GitHub](https://github.com)
2. 点击右上角的「+」→「New repository」
3. 仓库名称可以叫 `daily-paper-tracker` 或你喜欢的名字
4. 选择「Public」或「Private」（推荐 Private）
5. **不要**勾选「Initialize this repository with a README」
6. 点击「Create repository」

### 第二步：初始化本地 Git 仓库

在当前项目目录打开终端，执行以下命令：

```bash
# 初始化 Git
git init

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit"

# 关联远程仓库（替换为你的仓库地址）
git remote add origin https://github.com/你的用户名/你的仓库名.git

# 推送到 GitHub
git branch -M main
git push -u origin main
```

### 第三步：配置 GitHub Secrets

这是最关键的一步！

1. 进入你的 GitHub 仓库页面
2. 点击「Settings」（设置）
3. 在左侧菜单找到「Secrets and variables」→「Actions」
4. 点击「New repository secret」
5. 添加以下两个 Secret：

**第一个 Secret：**
- Name: `QQ_EMAIL`
- Value: `2026204614@qq.com`

**第二个 Secret：**
- Name: `QQ_PASSWORD`
- Value: 你的 QQ 邮箱授权码（不是登录密码！）

### 第四步：获取 QQ 邮箱授权码

如果你还没有授权码：

1. 登录 [QQ 邮箱](https://mail.qq.com)
2. 点击顶部「设置」→「账户」
3. 向下滚动找到「POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务」
4. 点击「开启」POP3/SMTP 服务
5. 按提示发送短信验证
6. 验证成功后会显示一个授权码，**复制保存好**
7. 将这个授权码填入 GitHub Secret 的 `QQ_PASSWORD`

### 第五步：测试工作流

1. 进入 GitHub 仓库的「Actions」标签页
2. 你会看到「Daily Paper Tracker」工作流
3. 点击「Run workflow」→「Run workflow」（绿色按钮）
4. 等待几分钟，看是否执行成功
5. 检查你的 QQ 邮箱是否收到邮件

### 第六步：确认定时任务

工作流已经配置为**每天 UTC 0 点（北京时间早上 8 点）**自动运行。

你可以在 `.github/workflows/daily-paper-tracker.yml` 中修改时间：

```yaml
schedule:
  - cron: '0 0 * * *'  # UTC 0点 = 北京时间 8点
```

Cron 格式：`分 时 日 月 周`

### 常见问题

**Q: 为什么邮件没收到？**
- 检查 GitHub Actions 执行日志是否报错
- 确认 Secrets 配置正确
- 检查 QQ 邮箱垃圾箱

**Q: 如何修改执行时间？**
- 修改 `.github/workflows/daily-paper-tracker.yml` 中的 cron 表达式
- 注意：GitHub Actions 使用 UTC 时间，北京时间 = UTC + 8

**Q: 论文和总结会保存到哪里？**
- 会自动提交到 GitHub 仓库的 `papers/` 目录
- 同时通过邮件发送给你

### 文件说明

```
每日论文讲解/
├── .github/
│   └── workflows/
│       └── daily-paper-tracker.yml  # GitHub Actions 配置
├── paper_tracker.py                  # 主脚本
├── requirements.txt                  # Python 依赖
├── .gitignore                        # Git 忽略文件
├── .env.example                      # 环境变量模板
├── GitHub设置指南.md                 # 本文件
└── 使用说明.md                       # 基本使用说明
```

## 祝你使用愉快！ 🎉

