import PyPDF2
import re
from collections import Counter
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from openpyxl import Workbook

def extract_text_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        num_pages = len(pdf_reader.pages)
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
            print(f"Extracting text from page {page_num + 1}/{num_pages}")
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

def main():
    pdf_file = r"D:\code\javascriptstudy\Andersen's Fairy Tales - Hans Christian Andersen.pdf"  # 替换为你的PDF文件路径
    excel_file = r"D:\code\javascriptstudy\output1.xlsx"  # Excel文件路径
    text = extract_text_from_pdf(pdf_file)
    cleaned_text = clean_text(text)
    high_freq_words = get_high_frequency_words(cleaned_text, num_words=100)
    print("Top 100 high frequency words:")
    for word, freq in high_freq_words:
        print(f"{word}: {freq}")
    save_to_excel(high_freq_words, excel_file)
    print(f"Results saved to {excel_file}")

if __name__ == "__main__":
    main()

