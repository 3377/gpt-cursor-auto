#!/usr/bin/env python3
"""
Microsoft邮件处理脚本
用于获取邮箱中的验证码
"""

import requests
import logging
from datetime import datetime
from typing import Dict, List
import configparser
import winreg
import time
import re

def get_proxy():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as key:
            proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
            
            if proxy_enable and proxy_server:
                proxy_parts = proxy_server.split(":")
                if len(proxy_parts) == 2:
                    return {"http": f"http://{proxy_server}", "https": f"http://{proxy_server}"}
    except WindowsError:
        pass
    return {"http": None, "https": None}

def load_config():
    """从config.txt加载配置"""
    config = configparser.ConfigParser()
    config.read('config.txt', encoding='utf-8')
    return config

def save_config(config):
    """保存配置到config.txt"""
    with open('config.txt', 'w', encoding='utf-8') as f:
        config.write(f)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

config = load_config()
microsoft_config = config['microsoft']

CLIENT_ID = microsoft_config['client_id']
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
TOKEN_URL = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'

class EmailClient:
    def __init__(self):
        config = load_config()
        if not config.has_section('tokens'):
            config.add_section('tokens')
        self.config = config
        self.refresh_token = config['tokens'].get('refresh_token', '')
        self.access_token = config['tokens'].get('access_token', '')
        expires_at_str = config['tokens'].get('expires_at', '1970-01-01 00:00:00')
        self.expires_at = datetime.strptime(expires_at_str, '%Y-%m-%d %H:%M:%S').timestamp()
    
    def is_token_expired(self) -> bool:
        """检查access token是否过期或即将过期"""
        buffer_time = 300
        return datetime.now().timestamp() + buffer_time >= self.expires_at
    
    def refresh_access_token(self) -> None:
        """刷新访问令牌"""
        refresh_params = {
            'client_id': CLIENT_ID,
            'refresh_token': self.refresh_token,
            'grant_type': 'refresh_token',
        }
        
        try:
            response = requests.post(TOKEN_URL, data=refresh_params, proxies=get_proxy())
            response.raise_for_status()
            tokens = response.json()
            
            self.access_token = tokens['access_token']
            self.expires_at = time.time() + tokens['expires_in']
            expires_at_str = datetime.fromtimestamp(self.expires_at).strftime('%Y-%m-%d %H:%M:%S')
            
            self.config['tokens']['access_token'] = self.access_token
            self.config['tokens']['expires_at'] = expires_at_str
            
            if 'refresh_token' in tokens:
                self.refresh_token = tokens['refresh_token']
                self.config['tokens']['refresh_token'] = self.refresh_token
            save_config(self.config)
        except requests.RequestException as e:
            logger.error(f"刷新访问令牌失败: {e}")
            raise

    def ensure_token_valid(self):
        """确保token有效"""
        if not self.access_token or self.is_token_expired():
            self.refresh_access_token()

    def get_latest_messages(self, top: int = 10) -> List[Dict]:
        """获取最新的邮件(包括收件箱和垃圾邮件)
        
        Args:
            top: 获取的邮件数量
            
        Returns:
            List[Dict]: 按时间倒序排列的邮件列表
        """
        self.ensure_token_valid()
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json',
            'Prefer': 'outlook.body-content-type="text"'
        }
        
        query_params = {
            '$top': top,
            '$select': 'subject,receivedDateTime,from,body,parentFolderId',
            '$orderby': 'receivedDateTime DESC',
            '$filter': "parentFolderId eq 'inbox' or parentFolderId eq 'junkemail'"
        }
        
        try:
            response = requests.get(
                f'{GRAPH_API_ENDPOINT}/me/messages',
                headers=headers,
                params=query_params,
                proxies=get_proxy()
            )
            response.raise_for_status()
            return response.json()['value']
        except requests.RequestException as e:
            logger.error(f"获取邮件失败: {e}")
            if response.status_code == 401:
                self.refresh_access_token()
                return self.get_latest_messages(top)
            raise

    def extract_verification_code(self, content: str, subject: str) -> str:
        """从邮件内容中提取验证码
        
        Args:
            content: 邮件内容
            subject: 邮件主题
            
        Returns:
            str: 提取到的验证码,如果没有找到则返回None
        """
        verify_subjects = ['verify', '验证', '验证码', 'code', '校验码']
        
        if not any(keyword.lower() in subject.lower() for keyword in verify_subjects):
            return None
            
        code_contexts = [
            r'验证码[是为:：]\s*(\d{6})',
            r'verification code[: ]?(\d{6})',
            r'code[: ]?(\d{6})',
            r'[\s:：](\d{6})[\s.]',
        ]
        
        for pattern in code_contexts:
            matches = re.findall(pattern, content, re.IGNORECASE)
            if matches:
                return matches[0]
                
        return None

    def get_latest_verification_codes(self, top: int = 1) -> List[str]:
        """获取最新邮件中的验证码
        
        Args:
            top: 获取的邮件数量
            
        Returns:
            List[str]: 验证码列表
        """
        result = []
        messages = self.get_latest_messages(top=top)
        
        for msg in messages:
            code = self.extract_verification_code(msg['body']['content'], msg['subject'])
            if code and code not in result:
                result.append(code)
        
        return result

    def delete_all_messages(self) -> bool:
        """删除所有邮件(包括收件箱和垃圾邮件)
        
        Returns:
            bool: 删除是否成功
        """
        self.ensure_token_valid()
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json'
        }
        
        try:
            query_params = {
                '$select': 'id',
                '$filter': "parentFolderId eq 'inbox' or parentFolderId eq 'junkemail'"
            }
            
            response = requests.get(
                f'{GRAPH_API_ENDPOINT}/me/messages',
                headers=headers,
                params=query_params,
                proxies=get_proxy()
            )
            response.raise_for_status()
            messages = response.json()['value']
            
            # 删除所有邮件
            for msg in messages:
                delete_response = requests.delete(
                    f'{GRAPH_API_ENDPOINT}/me/messages/{msg["id"]}',
                    headers=headers,
                    proxies=get_proxy()
                )
                delete_response.raise_for_status()
            
            logger.info(f"成功删除 {len(messages)} 封邮件")
            return True
            
        except requests.RequestException as e:
            logger.error(f"删除邮件失败: {e}")
            if response.status_code == 401:
                self.refresh_access_token()
                return self.delete_all_messages()
            return False

def main():
    try:
        client = EmailClient()
        
        verification_codes = client.get_latest_verification_codes(top=1)
        
        print("\n最新邮件中的验证码:")
        if verification_codes:
            for code in verification_codes:
                print(f"找到验证码: {code}")
                
            print("\n已获取到验证码，正在清理邮箱...")
            if client.delete_all_messages():
                print("所有邮件已清除")
            else:
                print("清除邮件失败")
        else:
            print("未找到验证码，不执行清理操作")
            
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        raise

if __name__ == '__main__':
    main()
