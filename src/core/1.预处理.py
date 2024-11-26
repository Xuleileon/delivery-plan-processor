import win32com.client
import os

try:
    # 1. 正确初始化 Excel 应用程序
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.DisplayAlerts = False  # 禁用警告弹窗
    
    # 2. 加载工作簿
    file_path = r'D:\010101010101\亚朵零售供应链\核心逻辑\表优化.xlsx'
    print(f"检查文件路径: {file_path}")
    print(f"文件是否存在: {os.path.exists(file_path)}")
    
    # 获取文件的绝对路径
    abs_path = os.path.abspath(file_path)
    workbook = excel.Workbooks.Open(abs_path)

    # 3. 计算常规产品和S级产品工作表的公式
    sheets_to_calculate = ['常规产品', 'S级产品']
    for sheet_name in sheets_to_calculate:
        sheet = workbook.Sheets(sheet_name)
        sheet.Calculate()  # 计算当前工作表的公式

        # 复制并粘贴值
        used_range = sheet.UsedRange
        used_range.Copy()  # 复制当前工作表的使用范围
        used_range.PasteSpecial(Paste=-4163)  # 粘贴值，-4163 表示 xlPasteValues

    # 4. 删除除了常规产品、S级产品和汇总之外的所有工作表
    sheets_to_keep = ['常规产品', 'S级产品', '汇总']
    # 从后向前删除工作表，避免索引变化带来的问题
    for i in range(workbook.Sheets.Count, 0, -1):
        sheet = workbook.Sheets(i)
        if sheet.Name not in sheets_to_keep:
            sheet.Delete()

    # 5. 保存修改后的工作簿
    output_path = r'D:\010101010101\亚朵零售供应链\核心逻辑\updated_表优化.xlsx'
    print(f"输出路径: {output_path}")
    workbook.SaveAs(os.path.abspath(output_path))

except Exception as e:
    print(f"发生错误: {str(e)}")
    
finally:
    try:
        # 6. 清理资源
        if 'workbook' in locals():
            workbook.Close(SaveChanges=False)
        if 'excel' in locals():
            excel.Quit()
    except:
        pass

print(f"处理完成，结果已保存到 {output_path}") 