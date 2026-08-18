import json
import os

log_path = r"C:\Users\quang\.gemini\antigravity-ide\brain\96f85643-e6e2-4978-935e-6c0842c38a46\.system_generated\logs\transcript_full.jsonl"
ocr_text = ""

with open(log_path, "r", encoding="utf-8") as f:
    for line in f:
        if "==Start of OCR for page 13==" in line:
            data = json.loads(line)
            content = data.get("content", "")
            if isinstance(content, list):
                for block in content:
                    if "text" in block and "==Start of OCR for page 13==" in block["text"]:
                        ocr_text = block["text"]
                        break
            elif isinstance(content, str):
                ocr_text = content
            break

if ocr_text:
    with open("ocr.txt", "w", encoding="utf-8") as f:
        f.write(ocr_text)
    print("Extracted OCR to ocr.txt")
else:
    print("OCR text not found in transcript!")
