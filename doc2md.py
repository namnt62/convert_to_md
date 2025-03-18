import os
import subprocess
import shutil
import pdfplumber
from docx import Document

# 用于处理 .doc → .docx 的库（需要安装）
import win32com.client as win32

# 设置源文件夹和目标文件夹路径
source_folder = r"D:\知识库"
target_folder = r"D:\new"

# 如果目标文件夹不存在，则创建
os.makedirs(target_folder, exist_ok=True)

# 将 .doc 转换为 .docx
def convert_doc_to_docx(doc_file):
    try:
        word = win32.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_file)
        docx_file = doc_file + "x"
        doc.SaveAs(docx_file, FileFormat=16)  # 16 = wdFormatDocumentDefault (docx)
        doc.Close()
        word.Quit()
        return docx_file
    except Exception as e:
        print(f"❌ Failed to convert {doc_file} to .docx: {e}")
        return None

# 将 PDF 转换为文本
def convert_pdf_to_text(pdf_file, text_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = '\n'.join([page.extract_text() or '' for page in pdf.pages])
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text)
        print(f"✅ Converted PDF to text: {pdf_file} → {text_file}")
        return text_file
    except Exception as e:
        print(f"❌ Failed to convert {pdf_file} to text: {e}")
        return None

# 遍历所有文件
for root, _, files in os.walk(source_folder):
    for filename in files:
        input_file = os.path.join(root, filename)
        output_file = os.path.join(target_folder, f"{os.path.splitext(filename)[0]}.md")
        
        # 1️⃣ 处理 .doc 文件
        if filename.endswith('.doc') and not filename.endswith('.docx'):
            print(f"🔄 Converting .doc → .docx: {input_file}")
            converted_file = convert_doc_to_docx(input_file)
            if converted_file:
                try:
                    subprocess.run(['markitdown', converted_file, '-o', output_file], check=True)
                    print(f"✅ Converted DOC to MD: {input_file} → {output_file}")
                    os.remove(converted_file)  # 清理中间文件
                except subprocess.CalledProcessError as e:
                    print(f"❌ Error converting DOCX: {e}")
            continue
        
        # 2️⃣ 处理 .docx 文件
        if filename.endswith('.docx'):
            try:
                subprocess.run(['markitdown', input_file, '-o', output_file], check=True)
                print(f"✅ Converted DOCX to MD: {input_file} → {output_file}")
            except subprocess.CalledProcessError as e:
                print(f"❌ Error converting DOCX: {e}")
            continue
        
        # 3️⃣ 处理 PDF 文件
        if filename.endswith('.pdf'):
            text_file = os.path.join(target_folder, f"{os.path.splitext(filename)[0]}.txt")
            if convert_pdf_to_text(input_file, text_file):
                try:
                    subprocess.run(['markitdown', text_file, '-o', output_file], check=True)
                    print(f"✅ Converted PDF to MD: {input_file} → {output_file}")
                    os.remove(text_file)  # 清理中间文件
                except subprocess.CalledProcessError as e:
                    print(f"❌ Error converting PDF to MD: {e}")
            continue
        
        # 4️⃣ 处理文本、HTML 等通用格式
        try:
            subprocess.run(['markitdown', input_file, '-o', output_file], check=True)
            print(f"✅ Converted {input_file} → {output_file}")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error converting {input_file}: {e}")
        except Exception as e:
            print(f"⚠️ Skipped {input_file}: {e}")

print("✅ 批量转换完成！")
