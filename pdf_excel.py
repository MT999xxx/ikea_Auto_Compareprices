import pdfplumber
import pandas as pd
import re
from pathlib import Path
import os
import logging

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def extract_order_info(pdf_path):
    """从PDF中提取订单信息"""
    order_items = []
    
    filename = os.path.basename(pdf_path)
    print(f"开始处理文件: {filename}")
    
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        
        # 提取订单号
        order_match = re.search(r'订单号[:：]\s*(\d+)', text)
        if not order_match:
            order_match = re.search(r'订单号.*?(\d{8,})', text)
        if not order_match:
            order_match = re.search(r'订单.*?号[：:]\s*(\d+)', text)
        if not order_match:
            # 尝试从文本中查找订单号格式的数字
            order_match = re.search(r'(?<!商品)(27\d{6})', text)
        
        if order_match:
            order_number = order_match.group(1)
            print(f"找到订单号: {order_number}")
        else:
            print("警告：未找到订单号")
            file_name = Path(pdf_path).stem
            # 尝试从文件名中提取订单号
            if file_name.startswith("CNREC"):
                order_number = file_name
            else:
                order_number = "未知"
        
        # 使用表格和文本结合的方式提取商品信息
        # 首先从表格中获取商品编号和金额
        items_from_table = extract_basic_items_from_table(page, order_number)
        
        # 然后从文本中确认商品数量并提取商品描述
        final_items = []
        for item in items_from_table:
            product_code = item['商品货号']
            
            # 在文本中精确查找这个商品的行
            qty = extract_accurate_quantity(text, product_code)
            
            # 提取商品描述
            description = extract_product_description(text, product_code)
            if description:
                item['商品名称与描述'] = description
            
            # 更新商品数量和单价
            if qty > 0:
                item['数量'] = qty
                # 重新计算单价
                item['商品单价'] = round(item['金额'] / qty, 2)
            
            final_items.append(item)
            print(f"最终商品信息: {product_code}, 数量: {item['数量']}, 单价: {item['商品单价']}, 金额: {item['金额']}, 描述: {item['商品名称与描述']}")
    
    return final_items

def extract_product_description(text, product_code):
    """从文本中提取商品描述"""
    description = ""
    
    # 在文本中查找包含该商品编号的行
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if product_code in line:
            # 找到包含商品编号的行
            
            # 完整的行文本
            full_line = line
            
            # 提取商品编号后面的字符串
            parts = full_line.split(product_code, 1)
            if len(parts) > 1:
                after_code = parts[1].strip()
                
                # 方法1: 提取英文单词和中文词组
                name_parts = []
                
                # 提取英文单词(全大写单词通常是产品名)
                eng_names = re.findall(r'([A-Z]{2,}(?:\s+[A-Z]+)*)', after_code)
                if eng_names:
                    name_parts.extend(eng_names)
                
                # 提取中文词组
                cn_names = re.findall(r'([\u4e00-\u9fa5]+(?:\s+[\u4e00-\u9fa5]+)*)', after_code)
                if cn_names:
                    name_parts.extend(cn_names)
                
                # 提取其它有意义的单词或短语
                other_words = re.findall(r'([a-zA-Z][a-zA-Z0-9]+)', after_code)
                for word in other_words:
                    if len(word) > 1 and word not in name_parts:
                        name_parts.append(word)
                
                # 提取尺寸信息
                size_info = re.findall(r'(\d+x\d+(?:x\d+)?)', after_code)
                if size_info:
                    name_parts.extend(size_info)
                
                # 组合描述 
                if name_parts:
                    description = " ".join(name_parts)
            
            # 如果描述仍为空，尝试查找下一行
            if not description and i + 1 < len(lines):
                next_line = lines[i + 1]
                if '¥' not in next_line and '%' not in next_line and not re.search(r'\d{3}\.\d{3}\.\d{2}', next_line):
                    # 提取有意义的词语
                    words = re.findall(r'([A-Z]{2,}|[\u4e00-\u9fa5]+|\d+x\d+(?:x\d+)?|[a-zA-Z][a-zA-Z0-9]+)', next_line)
                    if words:
                        description = " ".join(words)
            
            break
    
    # 通过特定格式从文本中查找商品名
    if not description:
        pattern = r'{}.*?([A-Z]+\s+[\u4e00-\u9fa5]+)'.format(re.escape(product_code))
        match = re.search(pattern, text)
        if match:
            description = match.group(1)
    
    # 清理描述文本
    if description:
        # 删除常见的无意义词语和数字
        description = re.sub(r'\b(件|个|元|米|厘米|套|组|盒|包|箱)\b', '', description)
        # 删除纯数字
        description = re.sub(r'\b\d+\b', '', description)
        # 清理多余空格
        description = re.sub(r'\s+', ' ', description).strip()
    
    return description

def extract_basic_items_from_table(page, order_number):
    """从表格中提取基本商品信息（商品编号、金额）"""
    items = []
    
    # 提取表格
    table = page.extract_table()
    if not table:
        return items
    
    # 查找商品编号列
    product_code_col = -1
    amount_col = -1
    name_col = -1  # 新增商品名称列识别
    
    # 首先查找表头
    for row_idx, row in enumerate(table):
        if row and any(cell and ('商品货号' in str(cell) or '订单号' in str(cell)) for cell in row):
            # 找到表头行，识别列
            for col_idx, cell in enumerate(row):
                if cell:
                    cell_text = str(cell).lower()
                    if '货号' in cell_text:
                        product_code_col = col_idx
                    elif '金额' in cell_text:
                        amount_col = col_idx
                    # 新增商品名称与描述列的识别
                    elif '名称' in cell_text or '描述' in cell_text:
                        name_col = col_idx
            
            # 如果找到了必要的列，从下一行开始处理商品
            if product_code_col != -1:
                for data_row in table[row_idx + 1:]:
                    if len(data_row) <= product_code_col:
                        continue
                        
                    product_code = data_row[product_code_col]
                    if not product_code or not isinstance(product_code, str):
                        continue
                        
                    product_code = product_code.strip()
                    
                    # 检查是否是有效的商品货号
                    if not re.match(r'\d{3}\.\d{3}\.\d{2}', product_code):
                        continue
                        
                    # 跳过自提/快递商品
                    if product_code.startswith('500.') and product_code != "500.005.97":
                        continue
                    
                    # 提取商品名称与描述（如果表格中有）
                    description = ""
                    if name_col != -1 and name_col < len(data_row) and data_row[name_col]:
                        description = str(data_row[name_col]).strip()
                    
                    # 提取金额
                    amount = 0
                    if amount_col != -1 and amount_col < len(data_row) and data_row[amount_col]:
                        amount_str = str(data_row[amount_col]).replace('¥', '').replace(',', '').strip()
                        try:
                            amount = float(amount_str)
                        except ValueError:
                            # 尝试提取数字部分
                            amount_match = re.search(r'([\d\.]+)', amount_str)
                            if amount_match:
                                amount = float(amount_match.group(1))
                    
                    # 添加商品（暂时使用默认数量1）
                    if product_code and amount > 0:
                        item = {
                            '订单号': order_number,
                            '商品货号': product_code,
                            '数量': 1,  # 默认值，稍后更新
                            '商品单价': amount,  # 默认等于金额，稍后更新
                            '现价': "",
                            '金额': amount,
                            '商品名称与描述': description  # 添加从表格中提取的描述
                        }
                        items.append(item)
            
            break
    
    return items

def extract_accurate_quantity(text, product_code):
    """精确提取商品数量"""
    # 默认数量
    quantity = 1
    
    # 查找包含该商品编号的行
    pattern = r'{}.*?(\d+)\s+[\d\.]+\s+(?:13|14)\s*%.*?¥'.format(re.escape(product_code))
    match = re.search(pattern, text)
    
    if match:
        # 找到匹配，提取数量
        quantity = int(match.group(1))
        return quantity
    
    # 尝试另一种模式：在商品编号后面紧跟着的数字可能是数量
    pattern = r'{}\s+.*?(?<!\d)(\d+)(?!\d).*?¥\s*[\d,\.]+'.format(re.escape(product_code))
    match = re.search(pattern, text)
    
    if match:
        # 验证提取的数字是否可能是数量（通常是1-10的小数字）
        qty = int(match.group(1))
        if 1 <= qty <= 20:  # 合理的数量范围
            quantity = qty
    
    # 使用更严格的模式查找
    pattern = r'{}.*?数量.*?(\d+)'.format(re.escape(product_code))
    match = re.search(pattern, text)
    
    if match:
        quantity = int(match.group(1))
    
    # 检查文本中是否有类似 "X件" 的模式
    pattern = r'{}.*?(\d+)\s*件'.format(re.escape(product_code))
    match = re.search(pattern, text)
    
    if match:
        quantity = int(match.group(1))
    
    # 查找商品编号和价格之间的明确数量标记
    content_after_code = text[text.find(product_code) + len(product_code):]
    qty_price_pattern = r'\s+(\d+)\s+\d+\.\d{2}\s+\d{2}\s*%'
    match = re.search(qty_price_pattern, content_after_code[:100])
    
    if match:
        qty = int(match.group(1))
        if qty <= 10:  # 合理的数量检查
            quantity = qty
    
    return quantity

def update_excel(items, excel_path):
    """更新Excel文件"""
    try:
        # 创建新数据的DataFrame
        df_new = pd.DataFrame(items)
        
        # 如果Excel文件存在，读取并合并数据
        if Path(excel_path).exists():
            df_existing = pd.read_excel(excel_path)
            
            # 确保两个DataFrame具有相同的列
            columns = ['订单号', '商品货号', '数量', '商品单价', '现价', '金额', '商品名称与描述']
            for col in columns:
                if col not in df_existing.columns:
                    df_existing[col] = ""
                if col not in df_new.columns:
                    df_new[col] = ""
                    
            # 准确选择需要的列并设置顺序
            df_existing = df_existing[columns]
            df_new = df_new[columns]
            
            # 合并数据前确保类型一致
            for col in ['数量', '商品单价', '金额']:
                if col in df_existing.columns and col in df_new.columns:
                    df_existing[col] = pd.to_numeric(df_existing[col], errors='coerce')
                    df_new[col] = pd.to_numeric(df_new[col], errors='coerce')
            
            # 合并数据
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new
        
        # 保存到Excel
        df_combined.to_excel(excel_path, index=False)
        print(f"成功保存数据到Excel文件: {excel_path}")
        
    except Exception as e:
        print(f"保存Excel文件时出错: {str(e)}")
        raise

def process_pdf_folder(pdf_folder, excel_path):
    """处理文件夹中的所有PDF文件"""
    # 确保PDF文件夹存在
    pdf_folder_path = Path(pdf_folder)
    if not pdf_folder_path.exists():
        print(f"错误：PDF文件夹不存在: {pdf_folder}")
        return
    
    # 获取所有PDF文件
    pdf_files = list(pdf_folder_path.glob("*.pdf"))
    if not pdf_files:
        print(f"警告：在 {pdf_folder} 中没有找到PDF文件")
        return
        
    print(f"找到 {len(pdf_files)} 个PDF文件")
    
    # 处理每个PDF文件
    success_count = 0
    error_count = 0
    all_items = []
    
    for pdf_file in pdf_files:
        try:
            print(f"\n开始处理文件: {pdf_file.name}")
            items = extract_order_info(str(pdf_file))
            if items:
                all_items.extend(items)
                print(f"成功处理文件: {pdf_file.name}，提取 {len(items)} 个商品")
                success_count += 1
            else:
                print(f"警告：从文件 {pdf_file.name} 中未提取到商品信息")
                error_count += 1
        except Exception as e:
            print(f"处理文件 {pdf_file.name} 时出错: {str(e)}")
            error_count += 1
            continue
    
    # 保存所有数据到Excel
    if all_items:
        update_excel(all_items, excel_path)
    
    # 打印处理结果统计
    print("\n处理完成！")
    print(f"成功处理: {success_count} 个文件")
    print(f"处理失败: {error_count} 个文件")
    print(f"总计文件: {len(pdf_files)} 个")
    print(f"总计提取: {len(all_items)} 个商品")

def main():
    # 设置路径
    pdf_folder = "F:\\宜家自动查询\\pdf"  # PDF文件夹路径
    excel_path = "F:\\宜家自动查询\\订单汇总.xlsx"  # Excel文件路径
    
    # 处理PDF文件夹
    process_pdf_folder(pdf_folder, excel_path)

if __name__ == "__main__":
    main()