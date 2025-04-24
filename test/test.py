import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import random
import logging
import re  # 将re模块移到全局导入
# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 设置用户代理
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0'
]

def get_product_url(product_number):
    """通过搜索获取商品的正确URL"""
    try:
        search_url = f"https://www.ikea.cn/cn/zh/search/products/?q={product_number}&qtype=search_keywords"
        
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Referer': 'https://www.ikea.cn/'
        }
        
        response = requests.get(search_url, headers=headers, timeout=10)
        response.raise_for_status()
        
        return search_url
        
    except Exception as e:
        logging.error(f"搜索商品 {product_number} 时出错: {str(e)}")
        return None



    """从宜家网站获取商品价格"""
    try:
        url = f"https://www.ikea.cn/cn/zh/p/mygglasvinge-mu-ge-si-wen-bei-tao-he-2-ge-zhen-tao-duo-se-{product_number}/"
        
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 更新价格选择器以匹配实际页面结构
        price_selectors = [
            '.pip-price',  # 新增
            '.product-pip-price',  # 新增
            '.pip-price-package',  # 新增
            '[data-product-price]',  # 新增
            'h2.pip-price-package__main-price',  # 新增
            'h3.pip-price-package__main-price',  # 新增
            '.pip-temp-price'  # 新增
        ]
        
        logging.info("开始查找价格元素...")
        for selector in price_selectors:
            elements = soup.select(selector)
            logging.info(f"选择器 {selector} 找到 {len(elements)} 个元素")
            
            for element in elements:
                try:
                    # 打印原始文本内容
                    text = element.get_text(strip=True)
                    logging.info(f"找到价格元素文本: {text}")
                    
                    # 使用正则表达式提取价格
                    import re
                    price_match = re.search(r'¥?\s*(\d+(?:\.\d{2})?)', text)
                    if price_match:
                        price = float(price_match.group(1))
                        if price > 0:
                            logging.info(f"成功解析价格: {price}")
                            return price
                except Exception as e:
                    logging.debug(f"处理价格元素时出错: {e}")
                    continue
        
        # 如果上述方法都失败，尝试在整个页面中查找价格
        price_patterns = [
            r'¥\s*(\d+(?:\.\d{2})?)',
            r'(\d+(?:\.\d{2})?)\s*元',
            r'price":\s*"?(\d+(?:\.\d{2})?)"?'  # 新增：查找JSON中的价格
        ]
        
        for pattern in price_patterns:
            matches = re.findall(pattern, response.text)
            if matches:
                try:
                    # 查找所有价格并返回最合理的一个
                    prices = [float(p) for p in matches if float(p) > 0]
                    if prices:
                        # 如果有多个价格，返回最常见的价格
                        from collections import Counter
                        price = Counter(prices).most_common(1)[0][0]
                        logging.info(f"通过页面文本匹配找到价格: {price}")
                        return price
                except ValueError:
                    continue
        
        logging.warning(f"无法获取商品 {product_number} 的价格信息")
        return None
            
    except Exception as e:
        logging.error(f"获取商品 {product_number} 价格时出错: {str(e)}")
        return None

def get_ikea_price(product_number):
    """从宜家网站获取商品价格"""
    try:
        url = f"https://www.ikea.cn/cn/zh/p/mygglasvinge-mu-ge-si-wen-bei-tao-he-2-ge-zhen-tao-duo-se-{product_number}/"
        
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 直接使用已确认有效的价格选择器
        price_element = soup.select_one('.price')
        
        if price_element:
            text = price_element.get_text(strip=True)
            logging.info(f"找到价格元素文本: {text}")
            
            # 提取价格
            price_match = re.search(r'¥?\s*(\d+(?:\.\d{2})?)', text)
            if price_match:
                price = float(price_match.group(1))
                logging.info(f"成功解析价格: {price}")
                return price
        
        # 如果上述方法失败，尝试在页面中查找价格
        price_patterns = [
            r'¥\s*(\d+(?:\.\d{2})?)',
            r'"price":\s*"?(\d+(?:\.\d{2})?)"?'
        ]
        
        for pattern in price_patterns:
            matches = re.findall(pattern, response.text)
            if matches:
                prices = [float(p) for p in matches if float(p) > 0]
                if prices:
                    price = min(prices)  # 使用最小价格，避免可能的促销价格干扰
                    logging.info(f"通过页面文本匹配找到价格: {price}")
                    return price
        
        logging.warning(f"无法获取商品 {product_number} 的价格信息")
        return None
            
    except Exception as e:
        logging.error(f"获取商品 {product_number} 价格时出错: {str(e)}")
        return None
    
def test_single_product(product_number):
    """测试单个商品的价格获取"""
    logging.info(f"测试商品 {product_number} 的价格获取")
    url = f"https://www.ikea.cn/cn/zh/p/mygglasvinge-mu-ge-si-wen-bei-tao-he-2-ge-zhen-tao-duo-se-{product_number}/"
    logging.info(f"访问URL: {url}")
    
    # 添加请求测试
    try:
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        }
        response = requests.get(url, headers=headers, timeout=10)
        logging.info(f"HTTP状态码: {response.status_code}")
        logging.info(f"响应头: {dict(response.headers)}")
    except Exception as e:
        logging.error(f"测试请求失败: {e}")
    
    price = get_ikea_price(product_number)
    logging.info(f"商品价格: {price}")
    
    return price

if __name__ == "__main__":
    # 设置更详细的日志级别
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # 测试单个商品
    test_product_number = "90591220"
    test_single_product(test_product_number)
    # 设置更详细的日志级别
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # 测试单个商品
    test_product_number = "90591220"  # 使用新的商品编号
    test_single_product(test_product_number)
    # 设置Excel文件路径和列名
    excel_file = "F:\\宜家自动查询\\test\\test1.xlsx"
    price_column = "价格"
    product_number_column = "货号"
    
    # 是否进行测试模式
    TEST_MODE = True
    
    if TEST_MODE:
        # 测试单个商品
        test_product_number = "70214286"  # 测试商品货号
        test_single_product(test_product_number)
    else:
        # 运行完整的价格监控
        monitor_prices(excel_file, price_column, product_number_column)