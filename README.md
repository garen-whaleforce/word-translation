# PDF to Traditional Chinese Word Translator

Upload an English PDF technical report, convert it to Word format, and translate the content to Traditional Chinese.

## Features

- PDF to DOCX conversion using LibreOffice
- English to Traditional Chinese translation using OpenAI GPT-4o
- Preserves document structure (paragraphs, tables)
- Simple web interface with drag-and-drop support

## Project Structure

```
word-translation/
├── app/
│   ├── __init__.py
│   ├── main.py           # FastAPI entry point
│   ├── routes.py         # API routes
│   ├── converter.py      # PDF → DOCX (LibreOffice)
│   └── docx_translate.py # DOCX translation logic
├── static/
│   ├── index.html        # Frontend page
│   ├── main.js           # Frontend JavaScript
│   └── style.css         # Styles
├── requirements.txt
├── Dockerfile
└── README.md
```

## Local Development

### Prerequisites

- Python 3.11+
- LibreOffice (for PDF to DOCX conversion)
- OpenAI API key

### Install LibreOffice

**macOS:**
```bash
brew install --cask libreoffice
```

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install libreoffice
```

### Setup

1. Clone the repository:
```bash
git clone <your-repo-url>
cd word-translation
```

2. Create and activate virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set environment variable:
```bash
export OPENAI_API_KEY="your-api-key-here"
```

5. Run the development server:
```bash
uvicorn app.main:app --reload
```

6. Open http://localhost:8000 in your browser.

## Zeabur Deployment

### Steps

1. Push your project to GitHub:
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin <your-github-repo-url>
git push -u origin main
```

2. Go to [Zeabur](https://zeabur.com) and create a new project.

3. Create a new service and select your GitHub repository.

4. Zeabur will automatically detect the Dockerfile and build the image.

5. Add environment variable in Zeabur:
   - Key: `OPENAI_API_KEY`
   - Value: Your OpenAI API key

6. Deploy and access your service via the provided URL.

## API Endpoints

### POST /api/upload

Upload a PDF file for translation.

**Request:**
- Content-Type: `multipart/form-data`
- Body: `file` - PDF file (max 20MB)

**Response:**
```json
{
  "file_id": "uuid",
  "download_url": "/api/download/<uuid>"
}
```

### GET /api/download/{file_id}

Download the translated Word document.

**Response:**
- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- File: `translated.docx`

### GET /api/healthz

Health check endpoint.

**Response:**
```json
{
  "status": "ok"
}
```

## Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `OPENAI_API_KEY` | OpenAI API key for translation | Yes |

## Limitations

- Maximum file size: 20MB
- Only PDF files are accepted
- LibreOffice conversion quality depends on the PDF structure
- Complex layouts may not be perfectly preserved

## License

MIT
