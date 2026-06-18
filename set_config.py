#获取textval 的数据,然后在这里打印出来

from pathlib import Path

# 模块级别的默认配置数据
config_data = {
    'product_name': '',
    'product_code': '',
    'calculation_unit': '1000',
    '月缴基数': '',
    '季缴基数': '',
    '半年缴基数': '',
    'premium_type': '保费',
}


def _load_product_codes() -> dict[str, str]:
    """从 productNameCode.txt 读取 产品名称→产品编码 映射。"""
    mapping: dict[str, str] = {}
    txt_path = Path(__file__).resolve().parent / "productNameCode.txt"
    if not txt_path.exists():
        return mapping
    for line in txt_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split("\t")
        if len(parts) >= 2:
            name = parts[0].strip()
            code = parts[1].strip()
            if name and code:
                mapping[name] = code
    return mapping


# 产品名称后缀，匹配时可能需要去掉
_NAME_SUFFIXES = ["费率表", "基本保险金额表", "_费率", "_基本保险金额"]


def _lookup_product(pdf_filename: str) -> tuple[str, str]:
    """根据 PDF 文件名查找产品名称和编码。

    匹配策略（按优先级）：
    1. 精确匹配 stem
    2. 去掉常见后缀后匹配
    3. 模糊包含匹配（stem 包含映射表中的名称）
    """
    stem = Path(pdf_filename).stem
    codes = _load_product_codes()

    if not codes:
        return stem, ""

    # 1. 精确匹配
    if stem in codes:
        return stem, codes[stem]

    # 2. 去后缀匹配
    cleaned = stem
    for suffix in _NAME_SUFFIXES:
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)]
            break
    if cleaned != stem and cleaned in codes:
        return cleaned, codes[cleaned]

    # 3. 模糊包含匹配
    for name, code in codes.items():
        if name in stem or stem in name:
            return name, code

    return stem, ""

def _normalize_text(text: str) -> str:
    text = text.replace('㌵', '月')
    text = text.replace('䍩', '费')
    text = text.replace('Ӕ', '交')
    text = text.replace('ᒤ', '年')
    text = text.replace('⌘', '注')
    text = text.replace('˖', '：')
    text = text.replace('؍', '保')
    text = text.replace('㑥', '保')
    text = text.replace('䰤', '间')
    text = text.replace('喴', '龄')
    text = text.replace('⭧', '男')
    text = text.replace('ྣ', '女')
    text = text.replace('ᴸ', '月')
    return text

def get_write_text(textval, pdf_filename=None) -> None:
    # textval数据 给get_content_vl.py 也页面数据用的
    # print('===text222=======', textval)
    
    import re
    textval = _normalize_text(textval)
    
    calculation_unit = '1000'  # 默认值
    
    # 定义匹配模式列表（按优先级排序）
    patterns = [
        r'每[0-9,，]+元保险费',
        r'每[0-9,，]+元年交保险费',
        r'每[0-9,，]+元基本保险金额',
        r'保险费：[0-9,，]+\s*元',
        r'每[0-9,，]+元趸交或年交保险费',
        r'每千元年交保险费',
        r'年交保险费：[0-9,，]+元',
        r'每万[元]?基本保险金额',  # 匹配 "每万元基本保险金额"
    ]
    
    # 遍历所有模式，找到第一个匹配
    for pattern in patterns:
        match = re.search(pattern, textval)
        if match:
            matched_text = match.group(0)
            print(f'匹配到: {matched_text}')
            
            # 从匹配的文本中提取数字
            number_match = re.search(r'[0-9,，]+', matched_text)
            if number_match:
                # 去除逗号并转换为整数
                number_str = number_match.group(0).replace(',', '').replace('，', '')
                calculation_unit = number_str
                print(f'calculation_unit={calculation_unit}')
                break
            elif '千' in matched_text:
                calculation_unit = '1000'
                print(f'calculation_unit={calculation_unit}')
                break
            elif '万' in matched_text:
                calculation_unit = '10000'
                print(f'calculation_unit={calculation_unit}')
                break
    else:
        # 如果没有匹配到任何模式，使用默认值
        print(f'未匹配到模式，使用默认值 calculation_unit={calculation_unit}')


    # payment_bases = {'月缴基数': '', '季缴基数': '', '半年缴基数': ''}  # 默认值
    payment_bases = {'月缴基数': '', '季缴基数': '', '半年缴基数': ''}  # 默认值
    
    # 匹配月缴基数
    month_patterns = [
        r'月交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保险费',
        r'月交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保费',
        r'月交保险费\s*[=＝]\s*年交保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'月交保费\s*[=＝]\s*年交保费\s*[*＊×xX]\s*([0-9.]+)',
        r'月交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保险费',
        r'月交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保费',
        r'月交保险费\s*[=＝]\s*年化保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'月交保费\s*[=＝]\s*年化保费\s*[*＊×xX]\s*([0-9.]+)',
        r'月交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额\s*÷\s*([0-9.]+)',
        r'月交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额([0-9.]+)',
    ]
    for pattern in month_patterns:
        match = re.search(pattern, textval)
        if match:
            payment_bases['月缴基数'] = match.group(1)
            print(f"月缴基数: {payment_bases['月缴基数']}")
            break

    # 匹配季缴基数
    quarter_patterns = [
        r'季交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保险费',
        r'季交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保费',
        r'季交保险费\s*[=＝]\s*年交保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'季交保费\s*[=＝]\s*年交保费\s*[*＊×xX]\s*([0-9.]+)',
        r'季交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保险费',
        r'季交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保费',
        r'季交保险费\s*[=＝]\s*年化保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'季交保费\s*[=＝]\s*年化保费\s*[*＊×xX]\s*([0-9.]+)',
        r'季交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额\s*÷\s*([0-9.]+)',
        r'季交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额([0-9.]+)',
    ]
    for pattern in quarter_patterns:
        match = re.search(pattern, textval)
        if match:
            payment_bases['季缴基数'] = match.group(1)
            print(f"季缴基数: {payment_bases['季缴基数']}")
            break

    # 匹配半年缴基数
    half_year_patterns = [
        r'半年交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保险费',
        r'半年交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年交保费',
        r'半年交保险费\s*[=＝]\s*年交保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'半年交保费\s*[=＝]\s*年交保费\s*[*＊×xX]\s*([0-9.]+)',
        r'半年交保险费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保险费',
        r'半年交保费\s*[=＝]\s*([0-9.]+)\s*[*＊×xX]\s*年化保费',
        r'半年交保险费\s*[=＝]\s*年化保险费\s*[*＊×xX]\s*([0-9.]+)',
        r'半年交保费\s*[=＝]\s*年化保费\s*[*＊×xX]\s*([0-9.]+)',
        r'半年交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额\s*÷\s*([0-9.]+)',
        r'半年交[0-9,，]+元保险费对应的基本保险金额[=＝]\s*年交[0-9,，]+元保险费对应的基本保险金额([0-9.]+)',
    ]
    for pattern in half_year_patterns:
        match = re.search(pattern, textval)
        if match:
            payment_bases['半年缴基数'] = match.group(1)
            print(f"半年缴基数: {payment_bases['半年缴基数']}")
            break
    
    print(f"payment_bases = {payment_bases}")

    if '保险金额对应的年交保险费' in textval:
        premium_type = '保额'
    else:
        premium_type = '保费'  # 默认，包含"保险费对应的基本保险金额"的情况

    # 从 PDF 文件名查找产品名称和编码
    if pdf_filename:
        product_name, product_code = _lookup_product(pdf_filename)
    else:
        product_name, product_code = '', ''

    # 将配置数据存储为模块级变量，供其他模块使用
    global config_data
    config_data = {
        'product_name': product_name,
        'product_code': product_code,  # 可以从其他地方获取
        'calculation_unit': calculation_unit,
        '月缴基数': payment_bases['月缴基数'],
        '季缴基数': payment_bases['季缴基数'],
        '半年缴基数': payment_bases['半年缴基数'],
        'premium_type': premium_type,
    }
    
    print(f"\n配置数据: {config_data}")
