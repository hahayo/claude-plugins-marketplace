"""
從大豐文件公版.docx 提取各 section 的 header/footer，建立乾淨的單頁範本。
使用「開新文件 + 複製 header/footer」的方式，避免 LinkToPrevious 繼承問題。
"""
import win32com.client, os, subprocess, time

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(2)

base = r"C:\Users\sophia\ClaudeProjects\行銷需求\brand\templates"
source = os.path.join(os.path.dirname(base), "Dafon文件素材", "大豐文件公版.docx")

# Section 對應表 (1-based index in 公版)
templates = {
    2: "template_橫式_中式logo.docx",
    3: "template_橫式_英式logo.docx",
    4: "template_直式_中式logo.docx",
    5: "template_直式_英式logo.docx",
}

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

src_doc = word.Documents.Open(os.path.abspath(source))

for sec_num, filename in templates.items():
    new_doc = word.Documents.Add()
    src_sec = src_doc.Sections(sec_num)
    new_sec = new_doc.Sections(1)

    # Copy page setup
    for attr in ['Orientation', 'PageWidth', 'PageHeight', 'TopMargin', 'BottomMargin',
                 'LeftMargin', 'RightMargin', 'HeaderDistance', 'FooterDistance']:
        setattr(new_sec.PageSetup, attr, getattr(src_sec.PageSetup, attr))

    # Copy header & footer
    src_sec.Headers(1).Range.Copy()
    new_sec.Headers(1).Range.Paste()
    src_sec.Footers(1).Range.Copy()
    new_sec.Footers(1).Range.Paste()

    out = os.path.join(base, filename)
    new_doc.SaveAs(os.path.abspath(out))
    new_doc.Close()
    print(f"OK: {filename}")

src_doc.Close(0)
word.Quit()
print("All templates generated.")
