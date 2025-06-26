from docx import Document
import sys
import os

def parse_docx_print(filename="test.docx"):
    if not os.path.exists(filename):
        print(f"파일이 존재 하지 않습니다: {filename}")
        return

    try:
        document=Document(filename)
        print(f"\n '{filename}' 문서 내용:\n" + "-"*50)

        for para in document.paragraphs:
            text=para.text.strip()
            if text:
                print(text)

    except Exception as e:
        print(f"파싱중 오류 발생: {e}")

if __name__=="__main__":
    parse_docx_print()