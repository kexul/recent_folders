#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建程序图标文件
只需要运行一次来生成图标文件
"""

from PIL import Image, ImageDraw
import os

def create_folder_icon():
    """创建文件夹图标"""
    # 创建一个简单的文件夹图标
    image = Image.new('RGB', (64, 64), color='white')
    draw = ImageDraw.Draw(image)
    
    # 绘制文件夹形状
    draw.rectangle([10, 20, 54, 50], fill='#FFD700', outline='#B8860B', width=2)
    draw.rectangle([10, 15, 25, 25], fill='#FFD700', outline='#B8860B', width=2)
    
    return image

def main():
    """生成图标文件"""
    # 创建图标
    icon = create_folder_icon()
    
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 保存为不同尺寸的PNG文件
    icon_32 = icon.resize((32, 32), Image.Resampling.LANCZOS)
    icon_16 = icon.resize((16, 16), Image.Resampling.LANCZOS)
    
    # 保存文件
    icon.save(os.path.join(current_dir, 'app_icon_64.png'))
    icon_32.save(os.path.join(current_dir, 'app_icon_32.png'))
    icon_16.save(os.path.join(current_dir, 'app_icon_16.png'))
    
    # 创建ICO文件（Windows标准图标格式）
    try:
        icon.save(os.path.join(current_dir, 'app_icon.ico'), 
                 sizes=[(16, 16), (32, 32), (64, 64)])
        print("图标文件创建成功：")
        print("- app_icon.ico (多尺寸ICO文件)")
        print("- app_icon_64.png (托盘图标)")
        print("- app_icon_32.png (窗口图标)")
        print("- app_icon_16.png (小图标)")
    except Exception as e:
        print(f"创建ICO文件失败: {e}")
        print("PNG文件已成功创建")

if __name__ == "__main__":
    main()