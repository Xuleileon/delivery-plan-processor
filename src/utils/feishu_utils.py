"""
飞书表格下载工具
"""
import os
import requests
from typing import Optional, List, Any, Dict, Tuple
import pandas as pd
from datetime import datetime
import string
import logging
import json

class FeishuSheetDownloader:
    """飞书表格下载器"""
    
    def __init__(self, app_id: Optional[str] = None, app_secret: Optional[str] = None):
        """
        初始化下载器
        
        Args:
            app_id: 飞书应用的 App ID
            app_secret: 飞书应用的 App Secret
        """
        self.app_id = app_id
        self.app_secret = app_secret
        self._tenant_access_token = None
        self.logger = logging.getLogger(__name__)
        
    def _get_tenant_access_token(self) -> str:
        """获取飞书租户访问令牌"""
        if not self.app_id or not self.app_secret:
            raise ValueError("请设置飞书应用的 App ID 和 App Secret")
            
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        headers = {"Content-Type": "application/json; charset=utf-8"}
        data = {
            "app_id": self.app_id,
            "app_secret": self.app_secret
        }
        
        try:
            self.logger.info(f"正在获取租户访问令牌，App ID: {self.app_id}")
            response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            
            result = response.json()
            if result.get("code") != 0:
                error_msg = f"获取租户访问令牌失败: {result}"
                self.logger.error(error_msg)
                self.logger.error(f"请求数据: {json.dumps(data, ensure_ascii=False)}")
                self.logger.error(f"响应内容: {json.dumps(result, ensure_ascii=False)}")
                raise Exception(error_msg)
                
            self.logger.info("成功获取租户访问令牌")
            return result["tenant_access_token"]
            
        except requests.exceptions.RequestException as e:
            error_msg = f"请求飞书API失败: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)

    def _get_sheet_data(self, spreadsheet_token: str, sheet_id: str) -> pd.DataFrame:
        """获取表格数据"""
        if not self._tenant_access_token:
            self._tenant_access_token = self._get_tenant_access_token()

        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}"
        headers = {
            "Authorization": f"Bearer {self._tenant_access_token}",
            "Content-Type": "application/json; charset=utf-8"
        }
        
        # 添加范围参数
        params = {
            "valueRenderOption": "ToString",  # 将所有值转换为字符串
            "dateTimeRenderOption": "FormattedString"  # 日期时间格式化为字符串
        }
        
        try:
            self.logger.info(f"正在获取表格数据，spreadsheet_token: {spreadsheet_token}, sheet_id: {sheet_id}")
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            
            result = response.json()
            if result.get("code") != 0:
                error_msg = f"获取表格数据失败: {result}"
                self.logger.error(error_msg)
                self.logger.error(f"请求URL: {url}")
                self.logger.error(f"响应内容: {json.dumps(result, ensure_ascii=False)}")
                raise Exception(error_msg)
                
            values = result["data"]["valueRange"]["values"]
            if not values:
                raise ValueError("表格数据为空")
                
            # 创建DataFrame并清理列名
            df = pd.DataFrame(values[1:], columns=values[0])
            df.columns = [str(col).strip() for col in df.columns]  # 确保列名是字符串并去除空白
            self.logger.info(f"成功获取表格数据，共 {len(df)} 行")
            return df
            
        except requests.exceptions.RequestException as e:
            error_msg = f"请求飞书API失败: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)

    def _get_sheet_metadata(self, spreadsheet_token: str) -> Dict[str, str]:
        """获取表格的元数据，包括sheet名称"""
        if not self._tenant_access_token:
            self._tenant_access_token = self._get_tenant_access_token()

        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo"
        headers = {
            "Authorization": f"Bearer {self._tenant_access_token}",
            "Content-Type": "application/json; charset=utf-8"
        }
        
        try:
            self.logger.info(f"正在获取表格元数据，spreadsheet_token: {spreadsheet_token}")
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            result = response.json()
            if result.get("code") != 0:
                error_msg = f"获取表格元数据失败: {result}"
                self.logger.error(error_msg)
                self.logger.error(f"请求URL: {url}")
                self.logger.error(f"响应内容: {json.dumps(result, ensure_ascii=False)}")
                raise Exception(error_msg)
                
            # 创建sheet_id到sheet_title的映射
            sheets = result["data"]["sheets"]
            sheet_names = {}
            for sheet in sheets:
                sheet_id = sheet.get("sheetId")
                sheet_title = sheet.get("title")
                if sheet_id and sheet_title:
                    sheet_names[sheet_id] = sheet_title
                    
            self.logger.info(f"成功获取表格元数据，包含 {len(sheet_names)} 个工作表")
            return sheet_names
            
        except requests.exceptions.RequestException as e:
            error_msg = f"请求飞书API失败: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)

    def download_sheets(self, sheet_urls: List[str], sheet_config: List[Dict[str, str]] = None) -> str:
        """
        下载多个表格并合并为一个Excel文件
        
        Args:
            sheet_urls: 飞书表格URL列表
            sheet_config: 工作表配置列表，每个配置包含 sheet_id 和 sheet_name
            
        Returns:
            保存的Excel文件路径
        """
        # 解析URL，获取spreadsheet_token和sheet_id
        processed_spreadsheets = {}
        
        for url in sheet_urls:
            self.logger.info(f"开始下载表格: {url}")
            try:
                # 从URL中提取token
                if "?" not in url:
                    spreadsheet_token = url.split("/")[-1]
                else:
                    spreadsheet_token = url.split("?")[0].split("/")[-1]
                
                # 获取表格元数据
                if spreadsheet_token not in processed_spreadsheets:
                    try:
                        processed_spreadsheets[spreadsheet_token] = self._get_sheet_metadata(spreadsheet_token)
                    except Exception as e:
                        self.logger.error(f"下载表格失败: {str(e)}")
                        # 创建一个空的DataFrame作为占位符
                        df = pd.DataFrame()
                        save_path = os.path.join(os.getcwd(), f'download_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
                        df.to_excel(save_path, sheet_name='Sheet1', index=False)
                        return save_path
                
            except Exception as e:
                self.logger.error(f"解析URL失败: {str(e)}")
                continue
        
        # 保存为Excel文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = os.path.join(os.getcwd(), f'download_{timestamp}.xlsx')
        
        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                if sheet_config:
                    # 使用配置的工作表信息
                    for sheet in sheet_config:
                        sheet_id = sheet['sheet_id']
                        sheet_name = sheet['sheet_name']
                        try:
                            df = self._get_sheet_data(spreadsheet_token, sheet_id)
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.logger.info(f"已下载工作表: {sheet_name}")
                        except Exception as e:
                            self.logger.error(f"下载工作表 {sheet_name} 失败: {str(e)}")
                            # 创建一个空的工作表作为占位符
                            pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # 使用元数据中的工作表信息
                    for spreadsheet_token, sheet_names in processed_spreadsheets.items():
                        for sheet_id, sheet_name in sheet_names.items():
                            try:
                                df = self._get_sheet_data(spreadsheet_token, sheet_id)
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                                self.logger.info(f"已下载工作表: {sheet_name}")
                            except Exception as e:
                                self.logger.error(f"下载工作表 {sheet_name} 失败: {str(e)}")
                                # 创建一个空的工作表作为占位符
                                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            self.logger.error(f"保存Excel文件失败: {str(e)}")
            # 确保至少有一个工作表
            df = pd.DataFrame()
            df.to_excel(save_path, sheet_name='Sheet1', index=False)
        
        return save_path

    def download_sheet(self, sheet_url: str, sheet_config: List[Dict[str, str]] = None) -> str:
        """
        下载单个飞书表格
        
        Args:
            sheet_url: 飞书表格的URL
            sheet_config: 工作表配置列表，每个配置包含 sheet_id 和 sheet_name
            
        Returns:
            保存的Excel文件路径
        """
        return self.download_sheets([sheet_url], sheet_config)
