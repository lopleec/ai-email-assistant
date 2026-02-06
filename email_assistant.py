#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI邮件自动回复助手
功能：自动读取Outlook邮箱的未读邮件，使用OpenAI兼容API生成智能回复并自动发送
"""

import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
import time
import os
from dotenv import load_dotenv
from openai import OpenAI
import logging
from datetime import datetime

# 加载环境变量
load_dotenv()

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_assistant.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class EmailAssistant:
    def __init__(self):
        """初始化邮件助手"""
        self.email_address = os.getenv('EMAIL_ADDRESS')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        self.openai_base_url = os.getenv('OPENAI_BASE_URL', 'https://api.openai.com/v1')
        self.check_interval = int(os.getenv('CHECK_INTERVAL', 60))
        self.ai_model = os.getenv('AI_MODEL', 'gpt-3.5-turbo')
        self.ai_temperature = float(os.getenv('AI_TEMPERATURE', 0.7))
        self.reply_language = os.getenv('REPLY_LANGUAGE', 'zh-CN')
        
        # Outlook IMAP和SMTP服务器
        self.imap_server = 'outlook.office365.com'
        self.smtp_server = 'smtp.office365.com'
        self.smtp_port = 587
        
        # 初始化OpenAI客户端（支持自定义base_url）
        self.openai_client = OpenAI(
            api_key=self.openai_api_key,
            base_url=self.openai_base_url
        )
        
        logger.info("邮件助手初始化完成")
        logger.info(f"使用OpenAI兼容API: {self.openai_base_url}")

    def decode_subject(self, subject):
        """解码邮件主题"""
        decoded_parts = decode_header(subject)
        subject_str = ''
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                subject_str += part.decode(encoding or 'utf-8', errors='ignore')
            else:
                subject_str += part
        return subject_str

    def get_email_body(self, msg):
        """获取邮件正文"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    try:
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or 'utf-8'
                        body = payload.decode(charset, errors='ignore')
                        break
                    except Exception as e:
                        logger.error(f"解析邮件正文出错: {e}")
        else:
            try:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or 'utf-8'
                body = payload.decode(charset, errors='ignore')
            except Exception as e:
                logger.error(f"解析邮件正文出错: {e}")
        return body.strip()

    def generate_ai_reply(self, sender, subject, body):
        """使用OpenAI生成邮件回复"""
        try:
            prompt = f"""你是一个专业的邮件助手。请根据以下邮件内容生成一封得体、专业的回复邮件。

发件人: {sender}
主题: {subject}
邮件内容:
{body}

请用{self.reply_language}语言生成回复。回复应该：
1. 礼貌且专业
2. 针对邮件内容给出恰当的回应
3. 语气自然友好
4. 长度适中（100-300字）

只返回回复的正文内容，不需要包含称呼和签名。"""

            response = self.openai_client.chat.completions.create(
                model=self.ai_model,
                messages=[
                    {"role": "system", "content": "你是一个专业的邮件助手，擅长撰写得体的邮件回复。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=self.ai_temperature,
                max_tokens=500
            )
            
            reply_content = response.choices[0].message.content.strip()
            logger.info(f"AI回复生成成功，长度: {len(reply_content)}")
            return reply_content
            
        except Exception as e:
            logger.error(f"AI生成回复失败: {e}")
            return "感谢您的来信。我已收到您的邮件，稍后会给您详细回复。"

    def send_reply(self, to_address, subject, body, original_msg_id=None):
        """发送回复邮件"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_address
            msg['To'] = to_address
            msg['Subject'] = f"Re: {subject}" if not subject.startswith('Re:') else subject
            
            if original_msg_id:
                msg['In-Reply-To'] = original_msg_id
                msg['References'] = original_msg_id
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_address, self.email_password)
                server.send_message(msg)
            
            logger.info(f"回复邮件已发送至: {to_address}")
            return True
            
        except Exception as e:
            logger.error(f"发送邮件失败: {e}")
            return False

    def process_unread_emails(self):
        """处理所有未读邮件"""
        try:
            # 连接IMAP服务器
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.email_address, self.email_password)
            mail.select('INBOX')
            
            # 搜索未读邮件
            status, messages = mail.search(None, 'UNSEEN')
            email_ids = messages[0].split()
            
            if not email_ids:
                logger.info("没有未读邮件")
                mail.logout()
                return
            
            logger.info(f"发现 {len(email_ids)} 封未读邮件")
            
            for email_id in email_ids:
                try:
                    # 获取邮件
                    status, msg_data = mail.fetch(email_id, '(RFC822)')
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # 解析邮件信息
                    sender = email.utils.parseaddr(msg['From'])[1]
                    subject = self.decode_subject(msg['Subject'] or '(无主题)')
                    msg_id = msg['Message-ID']
                    body = self.get_email_body(msg)
                    
                    logger.info(f"处理邮件 - 发件人: {sender}, 主题: {subject}")
                    
                    # 生成AI回复
                    reply_body = self.generate_ai_reply(sender, subject, body)
                    
                    # 发送回复
                    if self.send_reply(sender, subject, reply_body, msg_id):
                        logger.info(f"成功回复邮件: {sender}")
                    else:
                        logger.warning(f"回复邮件失败: {sender}")
                    
                    # 标记为已读
                    mail.store(email_id, '+FLAGS', '\\Seen')
                    
                    # 避免发送过快
                    time.sleep(2)
                    
                except Exception as e:
                    logger.error(f"处理邮件 {email_id} 时出错: {e}")
                    continue
            
            mail.logout()
            logger.info("邮件处理完成")
            
        except Exception as e:
            logger.error(f"连接邮箱或处理邮件时出错: {e}")

    def run(self):
        """启动邮件助手，持续监控"""
        logger.info(f"邮件助手启动，监控邮箱: {self.email_address}")
        logger.info(f"检查间隔: {self.check_interval}秒")
        
        while True:
            try:
                logger.info(f"--- 开始检查邮件 [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ---")
                self.process_unread_emails()
                logger.info(f"等待 {self.check_interval} 秒后再次检查...\n")
                time.sleep(self.check_interval)
                
            except KeyboardInterrupt:
                logger.info("\n邮件助手已停止")
                break
            except Exception as e:
                logger.error(f"运行时错误: {e}")
                logger.info("60秒后重试...")
                time.sleep(60)


def main():
    """主函数"""
    # 检查必要的环境变量
    required_vars = ['EMAIL_ADDRESS', 'EMAIL_PASSWORD', 'OPENAI_API_KEY']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        logger.error(f"缺少必要的环境变量: {', '.join(missing_vars)}")
        logger.error("请在 .env 文件中配置这些变量")
        return
    
    # 启动助手
    assistant = EmailAssistant()
    assistant.run()


if __name__ == '__main__':
    main()
