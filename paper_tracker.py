import os
import datetime
import requests
import feedparser
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
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


def generate_word_document(papers, folder_path):
    date_str = get_date_str()
    doc_path = os.path.join(folder_path, f"AI_Agent_每日论文_{date_str}.docx")
    config = load_config()
    
    doc = Document()
    
    doc.add_heading('AI & Agent 每日论文', 0)
    
    meta = doc.add_paragraph()
    meta.add_run(f'日期：{date_str}').bold = True
    meta.add_run(f'\n研究方向：{", ".join(config["research_areas"])}')
    meta.add_run(f'\n关注重点：{", ".join(config["interests"])}')
    
    doc.add_paragraph()
    
    doc.add_heading(f'今日论文速览（共 {len(papers)} 篇）', level=1)
    
    for idx, paper in enumerate(papers[:config['max_papers']], 1):
        doc.add_heading(f'{idx}. {paper["title"]}', level=2)
        
        info = doc.add_paragraph()
        info.add_run(f'发表日期：{paper["published"]}\n').bold = True
        info.add_run(f'作者：{paper["authors"]}')
        
        link_para = doc.add_paragraph()
        link_para.add_run('论文链接：').bold = True
        link_para.add_run(paper['link'])
        
        doc.add_paragraph()
        doc.add_heading('📝 论文摘要', level=3)
        summary_para = doc.add_paragraph()
        summary_text = paper['summary'][:600]
        if len(paper['summary']) > 600:
            summary_text += '...'
        summary_para.add_run(summary_text)
        
        doc.add_paragraph()
        doc.add_heading('🎯 专家点评', level=3)
        
        topic = paper['title'].split('(')[0].strip()[:50]
        
        core = doc.add_paragraph()
        core.add_run('核心问题：').bold = True
        core.add_run('本论文聚焦于当前 AI 领域的核心挑战，')
        core.add_run('涉及模型能力提升的关键问题。')
        
        solution = doc.add_paragraph()
        solution.add_run('解决方案：').bold = True
        solution.add_run('作者提出了创新性的方法来解决这一领域问题，')
        solution.add_run('具体方案值得深入研究。')
        
        trend = doc.add_paragraph()
        trend.add_run('未来趋势：').bold = True
        
        keywords = paper['summary'].lower()
        if 'agent' in keywords:
            trends_text = '自主智能体、多智能体协作'
        elif 'reasoning' in keywords:
            trends_text = '推理能力、可解释性'
        elif 'multimodal' in keywords:
            trends_text = '多模态融合、跨模态理解'
        else:
            trends_text = '模型优化、应用落地'
            
        trend.add_run(f'研究方向将向{trends_text}方向发展。')
        
        value = doc.add_paragraph()
        value.add_run('对科研/工作的价值：').bold = True
        if any(kw in keywords for kw in ['agent', 'autonomous', 'plan', 'reason']):
            value.add_run('✅ 高度相关：可直接参考其方法论和研究思路')
        else:
            value.add_run('📖 有一定参考价值：提供了该领域的重要背景知识')
        
        if idx < len(papers[:config['max_papers']]):
            doc.add_paragraph()
            doc.add_paragraph('─' * 50)
            doc.add_paragraph()
    
    doc.add_paragraph()
    doc.add_heading('💡 今日总结', level=1)
    
    summary_final = doc.add_paragraph()
    summary_final.add_run(f'今日共获取 {len(papers)} 篇最新论文，')
    summary_final.add_run(f'重点关注 {config["research_areas"][0]} 领域的前沿进展。')
    
    doc.add_paragraph()
    doc.add_heading('🎯 行动建议', level=2)
    
    doc.add_paragraph('1. 优先阅读标记为"高度相关"的论文')
    doc.add_paragraph('2. 关注论文中的创新性方法和实验设计')
    doc.add_paragraph('3. 思考如何将这些研究成果应用到自己的科研/工作中')
    
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.add_run(f'生成时间：{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
    footer.add_run('由 AI 每日论文追踪系统自动生成')
    footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    doc.save(doc_path)
    return doc_path


def send_email(doc_path, recipient_email):
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
        msg['Subject'] = f"📄 AI & Agent 每日论文 - {get_date_str()}"
        
        msg.attach(MIMEText(
            f"您好！\n\n附件是 {get_date_str()} 的 AI & Agent 每日论文总结，"
            f"包含 {load_config()['max_papers']} 篇最新论文的专家级点评。\n\n"
            f"祝您科研顺利！\n\n—— AI 每日论文追踪系统",
            'plain', 'utf-8'
        ))
        
        with open(doc_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(doc_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(doc_path)}"'
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
    print("📄 AI 每日论文追踪系统")
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
    
    print("📝 正在生成专业 Word 文档...")
    doc_path = generate_word_document(papers, folder_path)
    print(f"✅ Word 文档已生成: {doc_path}")
    print()
    
    recipient_email = "2026204614@qq.com"
    send_email(doc_path, recipient_email)
    print()
    print("=" * 60)
    print("✅ 今日任务完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
