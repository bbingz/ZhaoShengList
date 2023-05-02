import cv2
import torch
import numpy as np
from craft_text_detector import Craft
from PIL import Image

def main():
    # 加载预训练模型
    craft = Craft("craft_mlt_25k.pth")

    # 读取图像
    image = cv2.imread("img.png", cv2.IMREAD_UNCHANGED)

    # 检测文本区域
    prediction_result = craft.detect_text(image)

    # 绘制检测到的文本区域
    image_with_boxes = prediction_result.show(img=image)

    # 保存包含检测到的文本区域的图像
    cv2.imwrite("image_with_boxes.png", image_with_boxes)

    # 提取检测到的文本区域
    text_images = prediction_result.export(img=image)

    # 对检测到的文本区域应用 pytesseract 进行 OCR 识别
    recognized_texts = []
    for text_image in text_images:
        pil_text_image = Image.fromarray(text_image)
        recognized_text = pytesseract.image_to_string(pil_text_image, lang="chi_sim")
        recognized_texts.append(recognized_text.strip())

    # 打印识别出的文本
    print(recognized_texts)

if __name__ == "__main__":
    main()
