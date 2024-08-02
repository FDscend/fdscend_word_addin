import sys
import os
import re
import fitz  # PyMuPDF
from PIL import Image

def pdf_to_images(pdf_path, output_folder, zoom=3.0):
    # 打开 PDF 文件
    pdf_document = fitz.open(pdf_path)
    # 遍历 PDF 中的每一页
    for page_num in range(len(pdf_document)):
        # 获取页面
        page = pdf_document.load_page(page_num)
        # 设置缩放比例
        matrix = fitz.Matrix(zoom, zoom)
        # 渲染页面为图像
        pix = page.get_pixmap(matrix=matrix)
        # 将图像保存为文件
        image_path = f"{output_folder}/page_{page_num + 1}.png"
        pix.save(image_path)
        print(f"Page {page_num + 1} saved as {image_path}")


def merge_pdfs(pdf_folder, output_pdf):
    # 获取所有 PDF 文件
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
    pdf_files.sort()  # 确保按顺序处理文件

    # 创建一个新的空 PDF 文件
    merged_pdf = fitz.open()

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        with fitz.open(pdf_path) as pdf_document:
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                merged_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        print(f"Added {pdf_file}")

    # 保存合并后的 PDF
    merged_pdf.save(output_pdf)
    merged_pdf.close()
    print(f"PDF saved as {output_pdf}")


def natural_sort_key(s, _nsre=re.compile('([0-9]+)')):
    return [int(text) if text.isdigit() else text.lower() for text in _nsre.split(s)]

def images_to_pdf(image_folder, output_pdf):
    # 获取所有图片文件
    image_files = [f for f in os.listdir(image_folder) if f.endswith(('png', 'jpg', 'jpeg'))]
    image_files.sort(key=natural_sort_key)  # 按文件名中的数字排序

    # 创建一个新的空 PDF 文件
    pdf_document = fitz.open()

    for image_file in image_files:
        image_path = os.path.join(image_folder, image_file)
        img = Image.open(image_path)

        # Convert PIL image to RGB mode if not already
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        # Save the image as a temporary PDF
        img.save("temp.pdf", "PDF")

        # Insert the temporary PDF into the main PDF document
        pdf_page = fitz.open("temp.pdf")
        pdf_document.insert_pdf(pdf_page)
        print(f"Added {image_file}")

    # 保存合并后的 PDF
    pdf_document.save(output_pdf)
    pdf_document.close()
    print(f"PDF saved as {output_pdf}")
    os.remove("temp.pdf")


if __name__ == "__main__":

    match sys.argv[1]:
        case "pdf2img":
            pdf_path = sys.argv[2]
            output_folder = sys.argv[3]
            pdf_to_images(pdf_path, output_folder)
        case "mgpdf":
            pdf_folder = sys.argv[2]
            output_pdf = sys.argv[3]
            merge_pdfs(pdf_folder, output_pdf)
        case "img2pdf":
            image_folder = sys.argv[2]
            output_pdf = sys.argv[3]
            images_to_pdf(image_folder, output_pdf)
        case "pdf2imgpdf":
            pdf_path = sys.argv[2]
            output_pdf = sys.argv[3]
            tempDir = sys.argv[4]
            pdf_to_images(pdf_path, tempDir)
            images_to_pdf(tempDir, output_pdf)

