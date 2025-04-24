import requests
from bs4 import BeautifulSoup
import re
import random
import logging
import time

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 设置用户代理，模拟浏览器访问
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0'
]

def get_product_details(product_number):
    """获取商品详细信息，包括原价和促销价"""
    try:
        # 尝试直接访问商品详情页面
        urls = [
            f"https://www.ikea.cn/cn/zh/p/akern-a-ke-nei-li-jia-dian-tao-lan-se-xiu-hua-{product_number}/",
            f"https://www.ikea.cn/cn/zh/search/products/?q={product_number}&qtype=search_keywords"
        ]
        
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        }
        
        response = None
        for url in urls:
            try:
                logging.info(f"尝试URL: {url}")
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    logging.info(f"成功获取页面: {url}")
                    break
            except Exception as e:
                logging.warning(f"URL {url} 访问失败: {str(e)}")
        
        if not response or response.status_code != 200:
            logging.error(f"无法获取商品 {product_number} 页面")
            return {
                "product_number": product_number,
                "original_price": None,
                "current_price": None,
                "is_on_sale": False
            }
        
        # 保存HTML内容用于调试
        with open(f"ikea_{product_number}.html", "w", encoding="utf-8") as f:
            f.write(response.text)
        logging.info(f"已保存HTML内容到 ikea_{product_number}.html")
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 提取价格文本
        price_text = ""
        price_elements = soup.select('.price, .pip-price-package, .pip-price')
        for element in price_elements:
            price_text += element.get_text(strip=True)
            logging.info(f"价格元素文本: {element.get_text(strip=True)}")
        
        # 直接从页面文本中查找价格模式
        html_text = response.text
        
        # 查找促销价和原价
        original_price = None
        current_price = None
        
        # 优先查找特定模式
        if "会员价" in html_text or "非会员价" in html_text:
            # ÅKERNEJLIKA 阿克奈利加模式：非会员价¥129.00¥69.00
            match = re.search(r'非会员价\s*¥\s*(\d+(?:\.\d{2})?)\s*¥\s*(\d+(?:\.\d{2})?)', html_text)
            if match:
                original_price = float(match.group(1))
                current_price = float(match.group(2))
                logging.info(f"从会员价模式找到 - 原价: {original_price}, 现价: {current_price}")
        
        # 如果上面的模式没有匹配，尝试BRUNKRISSLA布朗瑞拉模式
        if not original_price or not current_price:
            # 查找模式：¥199.00¥149.00
            match = re.search(r'¥\s*(\d+(?:\.\d{2})?)\s*¥\s*(\d+(?:\.\d{2})?)', html_text)
            if match:
                # 第一个价格是原价，第二个是现价
                original_price = float(match.group(1))
                current_price = float(match.group(2))
                logging.info(f"从价格对模式找到 - 原价: {original_price}, 现价: {current_price}")
        
        # 如果仍然没有找到，尝试从页面中提取所有价格
        if not original_price or not current_price:
            all_prices = re.findall(r'¥\s*(\d+(?:\.\d{2})?)', html_text)
            if all_prices:
                prices = [float(p) for p in all_prices]
                unique_prices = sorted(set(prices))
                logging.info(f"页面中找到的所有价格: {unique_prices}")
                
                if len(unique_prices) >= 2:
                    # 假设最高的价格是原价，最低的是促销价
                    if unique_prices[0] < unique_prices[-1]:
                        current_price = unique_prices[0]
                        original_price = unique_prices[-1]
                        logging.info(f"从所有价格中选择 - 原价(最高价): {original_price}, 现价(最低价): {current_price}")
                elif len(unique_prices) == 1:
                    current_price = unique_prices[0]
                    original_price = unique_prices[0]
                    logging.info(f"只找到一个价格: {current_price}")
        
        # 分析促销信息
        is_on_sale = False
        if original_price and current_price and original_price > current_price:
            is_on_sale = True
        
        # 从页面标签判断是否促销
        sale_indicators = [
            "优惠有效期", "更低价格", "会员价", "限时", "促销", "特价"
        ]
        
        for indicator in sale_indicators:
            if indicator in html_text:
                logging.info(f"找到促销指标: {indicator}")
                if original_price and current_price and original_price > current_price:
                    is_on_sale = True
                    break
        
        return {
            "product_number": product_number,
            "original_price": original_price,
            "current_price": current_price,
            "is_on_sale": is_on_sale
        }
            
    except Exception as e:
        logging.error(f"获取商品 {product_number} 详细信息时出错: {str(e)}")
        return {
            "product_number": product_number,
            "original_price": None,
            "current_price": None,
            "is_on_sale": False
        }

if __name__ == "__main__":
    # 测试货号
    test_products = [
        "20571800",  # ÅKERNEJLIKA 阿克奈利加
        "90554802"   # BRUNKRISSLA 布朗瑞拉
    ]
    
    print("\n=== 开始测试商品价格获取 ===\n")
    
    results = []
    for product_number in test_products:
        print(f"\n正在测试商品 {product_number}...")
        details = get_product_details(product_number)
        results.append(details)
        print(f"商品: {product_number}")
        print(f"原价: {details['original_price']}")
        print(f"现价: {details['current_price']}")
        print(f"是否促销: {details['is_on_sale']}")
        time.sleep(2)  # 添加延迟，避免请求过快
    
    print("\n=== 测试结果汇总 ===")
    for result in results:
        print(f"商品 {result['product_number']}: 原价 ¥{result['original_price']:.2f}, 现价 ¥{result['current_price']:.2f}, 降价: {result['is_on_sale']}")