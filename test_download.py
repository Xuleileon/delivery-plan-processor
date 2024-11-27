"""
测试飞书表格下载
"""
from src.utils.feishu_utils import FeishuSheetDownloader

def main():
    # 创建下载器实例
    downloader = FeishuSheetDownloader(
        app_id="cli_a63afccc31bc900b",
        app_secret="hJLJHYk64H6nCSz3aq77ThJnJUzkOAC5"
    )
    
    # 要下载的表格URL列表
    sheet_urls = [
        "https://fr1r3d1ckr.feishu.cn/sheets/MdJWsy8N6hLqq7tCz51cPxySntf?sheet=RjyBZ8",
        "https://fr1r3d1ckr.feishu.cn/sheets/MdJWsy8N6hLqq7tCz51cPxySntf?sheet=64v01a",
        "https://fr1r3d1ckr.feishu.cn/sheets/MdJWsy8N6hLqq7tCz51cPxySntf?sheet=cxW3wO"
    ]
    
    try:
        saved_file = downloader.download_sheets(sheet_urls)
        print(f"\n所有表格已成功下载并合并到：{saved_file}")
    except Exception as e:
        print(f"\n下载失败：{str(e)}")

if __name__ == "__main__":
    main()
