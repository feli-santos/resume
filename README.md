# Resume — Felipe Santos

**Physical AI · Edge AI · AIoT Platforms · Solutions Architecture**

Tech Lead & Solutions Architect at Globant's Physical AI Studio. I build systems that connect
the physical and digital worlds — agentic AI platforms with LLM reasoning cores, real-time
computer vision at the edge, and IoT platforms managing fleets of 20,000+ devices.

📄 **[Download the PDF](./Felipe_Santos_Resume_Public.pdf)** · 📝 [DOCX](./Felipe_Santos_Resume_Public.docx)

---

## What's in here

| File | Description |
|------|-------------|
| `Felipe_Santos_Resume_Public.pdf` | Current resume, ATS-optimized, 2 pages |
| `Felipe_Santos_Resume_Public.docx` | Editable Word version |
| `build_resume.js` | Generates both files programmatically |

## Why a build script?

The resume is **generated from code**, not hand-edited in Word. One source of truth, consistent
typography, reproducible output, and a diffable history of every change.

```bash
npm install -g docx

node build_resume.js            # full version
node build_resume.js --public   # omits phone number
```

Export to PDF:

```bash
soffice --headless --convert-to pdf Felipe_Santos_Resume_Public.docx
```

### Design notes

- **Single column** — multi-column layouts break text extraction in most ATS parsers
- **Real text, no images** — everything is machine-readable
- **Semantic structure** — heading styles carry document outline, not just visual formatting
- **`keepNext` on subheadings** — prevents orphaned headings at page breaks

---

## Contact

[LinkedIn](https://www.linkedin.com/in/feli-santos/) · [Email](mailto:felipeaugustodosantos@gmail.com)
