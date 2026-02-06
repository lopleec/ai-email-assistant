#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OAuth2认证助手 - 用于Outlook邮箱登录
"""

import os
import json
import webbrowser
from msal import PublicClientApplication
import logging

logger = logging.getLogger(__name__)

# Microsoft Graph API 配置
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft官方开发者客户端ID
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = [
    "https://outlook.office365.com/IMAP.AccessAsUser.All",
    "https://outlook.office365.com/SMTP.Send"
]
TOKEN_FILE = "outlook_token.json"


class OutlookOAuthHelper:
    def __init__(self):
        """初始化OAuth助手"""
        self.app = PublicClientApplication(
            CLIENT_ID,
            authority=AUTHORITY
        )
        self.token_cache = None
        
    def load_token_cache(self):
        """从文件加载token缓存"""
        if os.path.exists(TOKEN_FILE):
            try:
                with open(TOKEN_FILE, 'r') as f:
                    self.token_cache = json.load(f)
                logger.info("已加载缓存的访问令牌")
                return True
            except Exception as e:
                logger.warning(f"加载token缓存失败: {e}")
        return False
    
    def save_token_cache(self, token_data):
        """保存token到文件"""
        try:
            with open(TOKEN_FILE, 'w') as f:
                json.dump(token_data, f, indent=2)
            logger.info("访问令牌已保存")
        except Exception as e:
            logger.error(f"保存token失败: {e}")
    
    def get_access_token(self, email_address):
        """获取访问令牌"""
        # 1. 尝试从缓存加载
        if self.load_token_cache():
            if self._is_token_valid():
                return self.token_cache.get('access_token')
        
        # 2. 尝试静默获取（使用刷新令牌）
        accounts = self.app.get_accounts()
        if accounts:
            logger.info("尝试使用刷新令牌静默获取访问令牌...")
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self.save_token_cache(result)
                return result['access_token']
        
        # 3. 需要用户交互式登录
        logger.info("需要重新登录，正在打开浏览器...")
        return self._interactive_login(email_address)
    
    def _interactive_login(self, email_address):
        """交互式登录（设备代码流）"""
        # 使用设备代码流，更适合CLI应用
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            logger.error("无法启动设备认证流程")
            return None
        
        print("\n" + "="*60)
        print("🔐 需要进行Outlook账户授权")
        print("="*60)
        print(f"\n请访问: {flow['verification_uri']}")
        print(f"并输入代码: {flow['user_code']}\n")
        print("然后使用你的Outlook账户登录并授权")
        print("="*60 + "\n")
        
        # 自动打开浏览器
        try:
            webbrowser.open(flow['verification_uri'])
        except:
            pass
        
        # 等待用户完成授权
        result = self.app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            logger.info("✅ 登录成功！")
            self.save_token_cache(result)
            return result['access_token']
        else:
            error = result.get("error_description", "未知错误")
            logger.error(f"❌ 登录失败: {error}")
            return None
    
    def _is_token_valid(self):
        """检查token是否有效"""
        if not self.token_cache:
            return False
        
        # 简单检查是否存在access_token
        return 'access_token' in self.token_cache


def authenticate_outlook(email_address):
    """认证Outlook账户并返回访问令牌"""
    helper = OutlookOAuthHelper()
    return helper.get_access_token(email_address)


if __name__ == '__main__':
    # 测试认证
    logging.basicConfig(level=logging.INFO)
    email = input("请输入你的Outlook邮箱地址: ")
    token = authenticate_outlook(email)
    if token:
        print(f"\n✅ 认证成功！")
        print(f"Access Token (前50字符): {token[:50]}...")
    else:
        print("\n❌ 认证失败")
