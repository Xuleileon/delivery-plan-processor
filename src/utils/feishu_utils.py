"""
飞书表格下载工具
"""
import os
import requests
from typing import Optional, List, Any, Dict, Tuple
import pandas as pd
from datetime import datetime
import string

class FeishuSheetDownloader:
    """飞书表格下载器"""
    
    def __init__(self, app_id: Optional[str] = None, app_secret: Optional[str] = None):
        """
        初始化下载器
        
        Args:
            app_id: 飞书应用的 App ID
            app_secret: 飞书应用的 App Secret
        """
        self.app_id = app_id or os.getenv("FEISHU_APP_ID")
        self.app_secret = app_secret or os.getenv("FEISHU_APP_SECRET")
        self._tenant_access_token = None
        
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
        
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        
        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"获取租户访问令牌失败: {result}")
            
        return result["tenant_access_token"]

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
        
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        
        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"获取表格数据失败: {result}")
            
        values = result["data"]["valueRange"]["values"]
        if not values:
            raise ValueError("表格数据为空")
            
        # 创建DataFrame并清理列名
        df = pd.DataFrame(values[1:], columns=values[0])
        df.columns = [str(col).strip() for col in df.columns]  # 确保列名是字符串并去除空白
        return df

    def _get_sheet_metadata(self, spreadsheet_token: str) -> Dict[str, str]:
        """获取表格的元数据，包括sheet名称"""
        if not self._tenant_access_token:
            self._tenant_access_token = self._get_tenant_access_token()

        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo"
        headers = {
            "Authorization": f"Bearer {self._tenant_access_token}",
            "Content-Type": "application/json; charset=utf-8"
        }
        
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"获取表格元数据失败: {result}")
            
        # 创建sheet_id到sheet_title的映射
        sheets = result["data"]["sheets"]
        sheet_names = {}
        for sheet in sheets:
            sheet_id = sheet.get("sheetId")
            sheet_title = sheet.get("title")
            if sheet_id and sheet_title:
                sheet_names[sheet_id] = sheet_title
        return sheet_names

    def download_sheets(self, sheet_urls: List[str], save_path: Optional[str] = None) -> str:
        """
        下载多个飞书表格并保存到一个Excel文件的不同sheet中
        
        Args:
            sheet_urls: 飞书表格的URL列表
            save_path: 保存路径，如果不指定则使用默认路径
            
        Returns:
            str: 保存的文件路径
        """
        # 创建Excel写入器
        if not save_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = os.path.join("output", f"feishu_sheets_{timestamp}.xlsx")
            
        # 确保输出目录存在
        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        
        # 使用ExcelWriter保存多个sheet
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            # 记录已处理的spreadsheet，避免重复获取元数据
            processed_spreadsheets = {}
            
            for url in sheet_urls:
                try:
                    print(f"\n开始下载表格: {url}")
                    spreadsheet_token = url.split("/")[-1].split("?")[0]
                    sheet_id = url.split("sheet=")[-1]
                    
                    # 获取sheet名称
                    if spreadsheet_token not in processed_spreadsheets:
                        processed_spreadsheets[spreadsheet_token] = self._get_sheet_metadata(spreadsheet_token)
                    sheet_names = processed_spreadsheets[spreadsheet_token]
                    sheet_title = sheet_names.get(sheet_id, sheet_id)  # 如果找不到名称，使用ID作为后备
                    
                    df = self._get_sheet_data(spreadsheet_token, sheet_id)
                    print(f"成功下载表格，共 {len(df)} 行数据")
                    
                    # 将数据保存到对应的sheet，使用原始sheet名称
                    df.to_excel(writer, sheet_name=sheet_title, index=False)
                    print(f"已保存到sheet: {sheet_title}")
                    
                except Exception as e:
                    print(f"下载表格失败: {str(e)}")
                    raise
        
        print(f"\n所有数据已保存到文件: {save_path}")
        return save_path
        
    def download_sheet(self, sheet_url: str, save_path: Optional[str] = None) -> str:
        """
        下载单个飞书表格
        
        Args:
            sheet_url: 飞书表格的URL
            save_path: 保存路径，如果不指定则使用默认路径
            
        Returns:
            str: 保存的文件路径
        """
        return self.download_sheets([sheet_url], save_path)
