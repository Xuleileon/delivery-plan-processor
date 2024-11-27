from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import uvicorn
from pathlib import Path
import tempfile
import shutil
import os
from typing import Optional

from main import process_delivery_plan

app = FastAPI(
    title="Delivery Plan Processor API",
    description="API for processing delivery plan Excel files",
    version="1.0.0"
)

@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    config_path: Optional[str] = None
):
    """
    处理上传的到货计划Excel文件
    
    Args:
        file: 上传的Excel文件
        config_path: 可选的配置文件路径
    
    Returns:
        处理结果，包含生成的文件路径
    """
    # 创建临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        # 保存上传的文件
        temp_input = Path(temp_dir) / file.filename
        with temp_input.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 处理文件
        result = process_delivery_plan(
            str(temp_input),
            output_dir=temp_dir,
            config_path=config_path
        )
        
        if not result['success']:
            raise HTTPException(status_code=400, detail=result['message'])
            
        # 将生成的文件移动到持久存储
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        final_files = {}
        for key, file_path in result['data'].items():
            src_path = Path(file_path)
            if src_path.exists():
                dest_path = output_dir / src_path.name
                shutil.copy2(src_path, dest_path)
                final_files[key] = str(dest_path)
        
        return {
            "success": True,
            "message": "文件处理成功",
            "files": final_files
        }

@app.get("/download/{filename}")
async def download_file(filename: str):
    """下载处理后的文件"""
    file_path = Path("output") / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(str(file_path))

if __name__ == "__main__":
    uvicorn.run("api_app:app", host="0.0.0.0", port=8000, reload=True)
