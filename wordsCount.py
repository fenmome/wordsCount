import PyPDF2
import re
from collections import Counter
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from openpyxl import Workbook
from tkinter import Tk, filedialog

def select_pdf_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return file_path

def select_excel_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return file_path

def extract_text_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        num_pages = len(pdf_reader.pages)
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
            print(f"正在从第 {page_num + 1} 页/共 {num_pages} 页提取文本")
    return text

def clean_text(text):
    # 去除非字母字符和空格
    cleaned_text = re.sub(r'[^a-zA-Z\s]', '', text)
    return cleaned_text

def get_high_frequency_words(text, num_words=100):
    # 分词
    words = word_tokenize(text.lower())
    # 去除停用词
    stop_words = set(stopwords.words('english'))
    filtered_words = [word for word in words if word not in stop_words]
    # 统计词频
    word_freq = Counter(filtered_words)
    # 获取频率最高的单词
    high_freq_words = word_freq.most_common(num_words)
    return high_freq_words

def save_to_excel(high_freq_words, excel_file):
    wb = Workbook()
    ws = wb.active
    ws.append(['Word', 'Frequency'])
    for word, freq in high_freq_words:
        ws.append([word, freq])
    wb.save(excel_file)

def get_num_words(max_attempts=3):
    """
    从用户处获取要提取的顶部单词的数量，并确保该输入是一个正整数。
    尝试次数限制为 max_attempts。
    """
    for attempt in range(max_attempts):
        print(f"尝试 {attempt+1}：")
        try:
            num_words = int(input("请输入要提取的顶部单词数量："))
            if num_words > 0:
                return num_words
            else:
                print("请输入一个正整数。")
        except ValueError:
            print("输入无效，请输入一个整数。")
    
    # 如果达到最大尝试次数仍未获取到有效输入
    if max_attempts > 0:
        print(f"达到最大尝试次数，未获取到有效输入。")
    return None  # 或者返回一个合理的默认值
def main():
    pdf_file = select_pdf_file()
    if not pdf_file:
        print("No file selected. Exiting...")
        return
    excel_file = select_excel_file()
    if not excel_file:
        print("No output file selected. Exiting...")
        return
    num_words = get_num_words()
    text = extract_text_from_pdf(pdf_file)
    cleaned_text = clean_text(text)
    high_freq_words = get_high_frequency_words(cleaned_text, num_words)
    print(f"Top {num_words} high frequency words:")
    for word, freq in high_freq_words:
        print(f"{word}: {freq}")
    save_to_excel(high_freq_words, excel_file)
    print(f"Results saved to {excel_file}")

if __name__ == "__main__":
    main()
