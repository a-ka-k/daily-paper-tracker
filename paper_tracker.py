import os
import datetime
import requests
import feedparser
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv


load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PAPERS_DIR = os.path.join(BASE_DIR, "papers")
CONFIG_FILE = os.path.join(BASE_DIR, "research_config.json")


def get_date_str():
    return datetime.datetime.now().strftime("%Y-%m-%d")


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        "research_areas": ["AI Agent", "Large Language Model", "Autonomous Systems"],
        "interests": ["科研", "工作"],
        "max_papers": 5
    }


def create_date_folder():
    date_str = get_date_str()
    folder_path = os.path.join(PAPERS_DIR, date_str)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def fetch_arxiv_papers(category="cs.AI", max_results=10):
    base_url = "http://export.arxiv.org/api/query"
    query = f"search_query=cat:{category}&sortBy=submittedDate&sortOrder=descending&max_results={max_results}"
    url = f"{base_url}?{query}"
    
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        feed = feedparser.parse(response.content)
        
        papers = []
        for entry in feed.entries:
            paper = {
                "title": entry.title.replace('\n', ' ').strip(),
                "authors": ", ".join([author.name for author in entry.authors[:5]]),
                "summary": entry.summary.replace('\n', ' ').strip(),
                "link": entry.link,
                "published": entry.published[:10]
            }
            papers.append(paper)
        return papers
    except Exception as e:
        print(f"获取论文失败: {e}")
        return []


def generate_expert_summary(papers, folder_path):
    date_str = get_date_str()
    summary_path = os.path.join(folder_path, f"{date_str}_AI进展总结.md")
    config = load_config()
    
    content = f"""# 🤖 AI & Agent 每日研究进展

**日期**: {date_str}  
**研究方向**: {' / '.join(config['research_areas'])}  
**关注重点**: {' / '.join(config['interests'])}

---

## 📊 今日论文速览

"""

    for idx, paper in enumerate(papers[:config['max_papers']], 1):
        content += f"""### {idx}. {paper['title']}

**📅 发表日期**: {paper['published']}  
**👥 作者**: {paper['authors']}  
**🔗 论文链接**: {paper['link']}

#### 📝 论文摘要
{paper['summary'][:500]}{'...' if len(paper['summary']) > 500 else ''}

#### 🎯 专家点评

本篇论文主要探讨了**{paper['title'].split('(')[0].strip() if '(' in paper['title'] else paper['title'][:30]}**相关领域的问题。

**核心问题**:
本论文聚焦于当前 AI 领域的核心挑战之一，涉及{'模型能力提升' if 'agent' in paper['title'].lower() else '技术创新'}的关键问题。

**解决方案**:
作者提出了创新性的方法来解决这一领域问题，具体方案值得深入研究。

**未来趋势**:
- 研究方向将向{'自主智能体' if 'agent' in paper['summary'].lower() else '模型优化'}方向发展
- 工程落地需要关注{'可解释性' if 'interpret' in paper['summary'].lower() else '实用性'}
- 学术界和工业界将更加重视{'安全性' if 'safe' in paper['summary'].lower() else '效果提升'}

**对科研/工作的价值**:
{'✅ 高度相关：可直接参考其方法论和研究思路' if any(keyword in paper['summary'].lower() for keyword in ['agent', 'autonomous', 'plan']) else '📖 有一定参考价值：提供了该领域的重要背景知识'}

---
"""

    content += f"""

## 💡 今日总结

今日共获取 **{len(papers)}** 篇最新论文，重点关注 **{config['research_areas'][0]}** 领域的前沿进展。

### 🎯 行动建议
1. 优先阅读标记为"高度相关"的论文
2. 关注论文中的创新性方法和实验设计
3. 思考如何将这些研究成果应用到自己的科研/工作中

---
*🤖 由 AI 每日论文追踪系统自动生成*  
*生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*

"""

    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(content)
    
    return summary_path


def send_email(summary_path, recipient_email):
    smtp_server = "smtp.qq.com"
    smtp_port = 587
    sender_email = os.getenv("QQ_EMAIL")
    sender_password = os.getenv("QQ_PASSWORD")
    
    if not sender_email or not sender_password:
        print("错误: 未配置 QQ_EMAIL 或 QQ_PASSWORD")
        return False
    
    print(f"正在发送邮件到: {recipient_email}")
    
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = f"🤖 AI & Agent 每日进展 - {get_date_str()}"
        
        with open(summary_path, "r", encoding="utf-8") as f:
            summary_content = f.read()
        
        msg.attach(MIMEText(summary_content, 'plain', 'utf-8'))
        
        with open(summary_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(summary_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(summary_path)}"'
            msg.attach(part)
        
        print("正在连接 SMTP 服务器...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        print("正在登录...")
        server.login(sender_email, sender_password)
        print("正在发送邮件...")
        server.send_message(msg)
        server.quit()
        print("✅ 邮件发送成功！")
        return True
    except Exception as e:
        print(f"❌ 邮件发送失败: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 60)
    print("🤖 AI 每日论文追踪系统")
    print("=" * 60)
    
    config = load_config()
    print(f"📌 当前研究方向: {', '.join(config['research_areas'])}")
    print(f"📌 关注重点: {', '.join(config['interests'])}")
    print()
    
    print("📥 正在获取最新论文...")
    folder_path = create_date_folder()
    papers = fetch_arxiv_papers()
    
    if not papers:
        print("❌ 未能获取到论文")
        return
    
    print(f"✅ 成功获取 {len(papers)} 篇论文")
    print()
    
    print("📝 正在生成专家级总结...")
    summary_path = generate_expert_summary(papers, folder_path)
    print(f"✅ 总结已生成: {summary_path}")
    print()
    
    recipient_email = "2026204614@qq.com"
    send_email(summary_path, recipient_email)
    print()
    print("=" * 60)
    print("✅ 今日任务完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
