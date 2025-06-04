import fitz  # PyMuPDF
from openai import OpenAI
import json
import re
from tkinter import filedialog, simpledialog, Tk

client = OpenAI(
    base_url="http://localhost:1234/v1",
    api_key="lm-studio"
)

def extract_text_from_pdf(pdf_path):
    print("📄 PDF 텍스트 추출 중...")
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    print("✅ 텍스트 추출 완료")
    return full_text

def generate_prompt_template_with_answers(full_text, num_items=5):
    print("🤖 LLM 프롬프트 생성 요청 중...")
    prompt = f'''
다음은 입력받은 문서서의 텍스트 내용입니다:

------------------------------
{full_text[:2000]}
------------------------------

이 문서의 내용을 바탕으로, 슬라이드 제작에 사용할 수 있도록 Q&A 목록 {num_items}개를 만들어줘.

각 항목은 아래 JSON 형식을 따르며, 'answer'는 슬라이드 목차 형태로 작성해줘.
특히 'answer' 항목은 실제 줄바꿈(엔터)을 포함한 다중 줄 텍스트로 작성해줘서, 사람이 보기 좋도록 해줘.

요구사항 요약:
- 출력은 반드시 JSON 배열 형식일 것
- 'answer'는 마크다운 구조로 작성할 것
- 모든 항목은 title, question, answer를 포함할 것

'''
    response = client.chat.completions.create(
        model="local-model",
        messages=[{"role": "user", "content": prompt}]
    )
    print("✅ 프롬프트 생성 완료")
    return response.choices[0].message.content

def extract_json_from_response(response):
    match = re.search(r"\[.*\]", response, re.DOTALL)
    if match:
        return match.group(0)
    return response

def fix_common_json_errors(json_str):
    json_str = json_str.replace("“", "\"").replace("”", "\"").replace("‘", "'").replace("’", "'")

    # "question": "..."\n"answer": "..." → 쉼표 추가
    json_str = re.sub(r'("question"\s*:\s*".+?")\s*("answer"\s*:\s*")', r'\1,\n\2', json_str, flags=re.DOTALL)

    # answer 내부 줄바꿈 문자열 정리
    def escape_answer(m):
        body = m.group(2).replace('\r', '').replace('\n', '\\n').replace('\t', '').strip()
        return f'{m.group(1)}{body}{m.group(3)}'

    json_str = re.sub(r'("answer"\s*:\s*")((?:[^"\\]|\\.)*?)(")', escape_answer, json_str, flags=re.DOTALL)

    # 연속된 객체 사이에 쉼표가 없을 경우 추가
    json_str = re.sub(r'}\s*{', '}, {', json_str)

    return json_str

def save_prompt_template_to_file(template_str, filename):
    print("💾 프롬프트 저장 중...")
    template_str = extract_json_from_response(template_str)

    try:
        templates = json.loads(template_str)
    except json.JSONDecodeError as e:
        print("⚠️ JSON 파싱 실패. 자동 수정을 시도합니다...")
        print(f"⛔ 원본 오류: {e}")
        print("📝 원본 응답 내용:")
        print(template_str)

        cleaned_str = fix_common_json_errors(template_str)
        try:
            templates = json.loads(cleaned_str)
            print("✅ 자동 수정 성공")
        except json.JSONDecodeError as e2:
            print("❌ 자동 수정 실패. 아래 응답을 수동으로 확인하세요:")
            print(cleaned_str)
            raise Exception(f"자동 수정을 실패했습니다: {str(e2)}")

    with open(filename, "w", encoding="utf-8") as f:
        f.write("prompt_templates = ")
        f.write(json.dumps(templates, indent=4, ensure_ascii=False))
    print(f"✅ 저장 완료: {filename}")

def main():
    root = Tk()
    root.withdraw()

    print("📂 PDF 파일 선택 중...")
    pdf_path = filedialog.askopenfilename(title="PDF 선택", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        print("❌ 작업 취소됨.")
        return

    num_items = simpledialog.askinteger("프롬프트 수", "몇 개의 프롬프트를 생성할까요?", minvalue=1, maxvalue=20)
    if not num_items:
        print("❌ 작업 취소됨.")
        return

    save_path = filedialog.asksaveasfilename(
        title="저장할 파일 이름",
        defaultextension=".py",
        filetypes=[("Python Files", "*.py")],
        initialfile="generated_prompt_templates.py"
    )
    if not save_path:
        print("❌ 작업 취소됨.")
        return

    text = extract_text_from_pdf(pdf_path)
    response = generate_prompt_template_with_answers(text, num_items=num_items)
    save_prompt_template_to_file(response, save_path)

    print("🎉 모든 작업이 완료되었습니다.")

if __name__ == "__main__":
    main()
