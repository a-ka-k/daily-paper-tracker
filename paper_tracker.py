import os
import datetime
import requests
import feedparser
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from bs4 import BeautifulSoup
from dotenv import load_dotenv


load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PAPERS_DIR = os.path.join(BASE_DIR, "papers")


def get_date_str():
    return datetime.datetime.now().strftime("%Y-%m-%d")


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
                "title": entry.title,
                "authors": ", ".join([author.name for author in entry.authors]),
                "summary": entry.summary,
                "link": entry.link,
                "pdf_link": entry.link.replace("abs", "pdf") + ".pdf",
                "published": entry.published
            }
            papers.append(paper)
        return papers
    except Exception as e:
        print(f"Error fetching arXiv papers: {e}")
        return []


def download_paper(pdf_url, save_path):
    try:
        response = requests.get(pdf_url, timeout=60)
        response.raise_for_status()
        with open(save_path, "wb") as f:
            f.write(response.content)
        return True
    except Exception as e:
        print(f"Error downloading paper: {e}")
        return False


def generate_summary(papers, folder_path):
    date_str = get_date_str()
    summary_path = os.path.join(folder_path, f"{date_str}_AI进展总结.md")
    
    content = f"# AI & Agent 研究进展 - {date_str}\n\n"
    content += "## 📄 最新论文\n\n"
    
    downloaded_files = []
    
    for idx, paper in enumerate(papers[:5], 1):
        content += f"### {idx}. {paper['title']}\n"
        content += f"- **作者**: {paper['authors']}\n"
        content += f"- **发布日期**: {paper['published']}\n"
        content += f"- **链接**: {paper['link']}\n"
        content += f"- **摘要**:\n{paper['summary']}\n\n"
        
        filename = f"{idx}_{paper['title'][:50].replace(' ', '_').replace('/', '_')}.pdf"
        save_path = os.path.join(folder_path, filename)
        if download_paper(paper['pdf_link'], save_path):
            downloaded_files.append(save_path)
            content += f"- 📥 已下载: {filename}\n\n"
    
    content += "---\n*由每日论文追踪系统自动生成*\n"
    
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(content)
    
    return summary_path, downloaded_files


def send_email(summary_path, downloaded_files, recipient_email):
    smtp_server = "smtp.qq.com"
    smtp_port = 587
    sender_email = os.getenv("QQ_EMAIL")
    sender_password = os.getenv("QQ_PASSWORD")
    
    if not sender_email or not sender_password:
        print("错误: 未配置 QQ_EMAIL 或 QQ_PASSWORD")
        return False
    
    print(f"正在发送邮件到: {recipient_email}")
    print(f"发件人: {sender_email}")
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"AI & Agent 每日进展 - {get_date_str()}"
    
    try:
        with open(summary_path, "r", encoding="utf-8") as f:
            summary_content = f.read()
        
        msg.attach(MIMEText(summary_content, 'plain', 'utf-8'))
        
        with open(summary_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(summary_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(summary_path)}"'
            msg.attach(part)
        
        for file_path in downloaded_files:
            try:
                with open(file_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                    msg.attach(part)
            except Exception as e:
                print(f"跳过附件 {file_path}: {e}")
        
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
    print("开始获取最新论文...")
    
    folder_path = create_date_folder()
    papers = fetch_arxiv_papers()
    
    if not papers:
        print("未能获取到论文")
        return
    
    print(f"获取到 {len(papers)} 篇论文")
    
    summary_path, downloaded_files = generate_summary(papers, folder_path)
    print(f"总结已生成: {summary_path}")
    
    recipient_email = "2026204614@qq.com"
    send_email(summary_path, downloaded_files, recipient_email)


if __name__ == "__main__":
    main()

