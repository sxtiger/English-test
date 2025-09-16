import pandas as pd

def markdown_to_excel(md_file, excel_file):
    """Markdown转Excel工具函数"""
    current_book = ""
    current_day = ""
    data = []
    
    try:
        with open(md_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                if line.startswith('# '):
                    current_book = line[2:].strip()
                    current_day = ""
                elif line.startswith('## '):
                    current_day = line[3:].strip()
                elif line.startswith('- '):
                    parts = line[2:].split(':', 1)
                    if len(parts) == 2:
                        data.append({
                            "英文短语": parts[0].strip(),
                            "中文翻译": parts[1].strip(),
                            "学习日": current_day,
                            "书本名称": current_book
                        })
        
        df = pd.DataFrame(data)
        df = df[['英文短语', '中文翻译', '学习日', '书本名称']]
        df.to_excel(excel_file, index=False)
        return len(df)
    except Exception as e:
        print(f"[ERROR] 文件转换失败: {str(e)}")
        raise