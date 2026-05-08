# 📄 → 🎯 Doc to PPT Converter

A simple and powerful Python tool that converts **DOCX and PDF** files into **PowerPoint presentations** automatically — no manual copy-pasting, no formatting headaches.

Built with Python + Streamlit. Upload your file, choose how to split it into slides, and download your PPTX in seconds.

---

## 🚀 Live Demo

👉 [Try it here](https://your-streamlit-app-link.streamlit.app) *(replace with your deployed link)*

---

## ✨ Features

- 📂 Supports both **DOCX** and **PDF** files
- 🔀 **4 ways to split** your document into slides:
  - **By Headings** — each heading becomes a new slide
  - **By Page Breaks** — each page becomes a slide
  - **By Custom Keyword** — you choose the keyword (e.g. Q, Section, Chapter)
  - **By Paragraphs** — groups every 3 paragraphs into a slide
- 🖼️ **Auto-extracts images** from the source file and places them on slides
- 🎨 **Dark theme styling** applied automatically — clean, professional look
- ⬇️ One-click **PPTX download**

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| Python | Core logic |
| Streamlit | Web UI |
| Pandoc | Document parsing & PPTX conversion |
| PyMuPDF (fitz) | PDF text & image extraction |
| python-docx | DOCX text & image extraction |
| python-pptx | PPTX styling |

---

## ⚙️ Run Locally

### 1. Clone the repo
```bash
git clone https://github.com/your-username/doc-to-ppt.git
cd doc-to-ppt
```

### 2. Install dependencies
```bash
pip install streamlit python-pptx python-docx pymupdf
```

> Also install **Pandoc**: https://pandoc.org/installing.html

### 3. Run the app
```bash
streamlit run doc_to_ppt_generic.py
```

---

## 📸 Screenshot

*(Add a screenshot of your Streamlit UI here)*

---

## 📌 Notes

- For **math-heavy PDFs** where text appears black on dark background: go to **View → Slide Master → Select text → Change font color → Close Slide Master**
- Custom keyword split works best for structured documents like Q&A sheets, reports with chapters, or sectioned notes

---

## 🙋‍♂️ Author

**Madhura Shirsikar**  
[LinkedIn](https://www.linkedin.com/in/your-profile) • [GitHub](https://github.com/your-username)

---

## 📃 License

MIT License — free to use and modify.
