import os
import sys
import uuid
import shutil
from typing import Dict, Any, List
from fastapi import FastAPI, HTTPException, BackgroundTasks, UploadFile, File, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn

# 动态添加路径以导入现有的解析器和生成器
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(os.path.join(parent_dir, 'docx解析器'))
sys.path.append(os.path.join(parent_dir, 'docx生成器'))

from enhanced_docx_parser import EnhancedDocxParser
from json_to_docx import JsonToDocxConverter

app = FastAPI(title="Docx 智能处理 API", description="支持 Docx 解析为 JSON 和从 JSON 还原 Docx")

# 配置存储目录
UPLOAD_DIR = os.path.join(current_dir, "uploads")
OUTPUT_DIR = os.path.join(current_dir, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 挂载静态文件目录，用于通过 URL 访问生成的文档
app.mount("/download", StaticFiles(directory=OUTPUT_DIR), name="download")

class ParseRequest(BaseModel):
    url: str

class GenerateRequest(BaseModel):
    json_data: Dict[str, Any]

def cleanup_file(path: str):
    """后台任务：清理临时文件"""
    if os.path.exists(path):
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception as e:
            print(f"清理文件失败: {e}")

@app.post("/parse_file")
async def parse_docx_file(file: UploadFile = File(...)):
    """
    解析接口：上传 Word 文件，直接返回生成的 JSON 内容
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="仅支持 .docx 格式文件")
    
    file_id = str(uuid.uuid4())
    temp_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
    
    parser = EnhancedDocxParser()
    try:
        # 1. 保存上传的文件
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 2. 提取样式
        parser.extract_styles_from_xml(temp_path)
        
        # 3. 解析内容
        result = parser.parse(temp_path)
        
        return result
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"文件解析失败: {str(e)}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        parser.cleanup()

@app.post("/parse")
async def parse_docx(request: ParseRequest):
    """
    解析接口：输入 Docx URL，返回结构化 JSON
    """
    parser = EnhancedDocxParser()
    try:
        # 1. 下载文件
        local_path = parser.download_docx(request.url)
        
        # 2. 提取样式
        parser.extract_styles_from_xml(local_path)
        
        # 3. 解析内容
        result = parser.parse(local_path)
        
        # 4. 清理临时下载的文件
        if os.path.exists(local_path):
            os.remove(local_path)
            
        return result
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"解析失败: {str(e)}")
    finally:
        parser.cleanup()

@app.post("/generate")
async def generate_docx(request: Request, data: GenerateRequest):
    """
    生成接口：输入 JSON 数据，返回生成的 Word 文档 URL
    """
    converter = JsonToDocxConverter()
    file_id = str(uuid.uuid4())
    output_filename = f"generated_{file_id}.docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    
    try:
        # 1. 转换并保存
        converter.convert(data.json_data, output_path)
        
        # 2. 构造访问 URL (强制使用 https)
        base_url = str(request.base_url).rstrip('/')
        if base_url.startswith("http://"):
            base_url = base_url.replace("http://", "https://", 1)
        download_url = f"{base_url}/download/{output_filename}"
        
        return {
            "message": "生成成功",
            "url": download_url,
            "filename": output_filename
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"生成失败: {str(e)}")

@app.get("/")
async def root():
    return {
        "message": "Docx Processing Service is running", 
        "endpoints": {
            "word_to_json": "/parse_file (POST, upload file)",
            "remote_word_to_json": "/parse (POST, json with url)",
            "json_to_word": "/generate (POST, json with data)"
        }
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8080)
