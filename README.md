# 🎯 SuuSri AI PPT Generator

<div align="center">

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![AI Powered](https://img.shields.io/badge/AI%20Powered-FF6B6B?style=for-the-badge&logo=openai&logoColor=white)

**An AI-Powered Presentation Generator - Transform Topics into Professional PowerPoint Presentations in Seconds**

Powered by LLMs | Built with Python & Flask | Deploy Anywhere

[Live Demo](#-live-demo) • [Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Tech Stack](#-tech-stack)

</div>

---

## 📋 Table of Contents

- [About](#-about)
- [Live Demo](#-live-demo)
- [Features](#-features)
- [Tech Stack](#-tech-stack)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Usage](#-usage)
- [Project Structure](#-project-structure)
- [API Documentation](#-api-documentation)
- [Deployment](#-deployment)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🎯 About

**SuuSri AI PPT Generator** is a revolutionary application that leverages cutting-edge AI technology to automatically generate professional PowerPoint presentations. Simply provide a topic or custom prompt, and the AI creates a complete, well-structured presentation with:

- 📊 Multiple slides with engaging content
- 🎨 Professional formatting and layouts
- 📝 Coherent, well-organized information
- ⚡ Instant generation (seconds, not hours)
- 📥 Ready-to-download PowerPoint files
- 🔄 Fully customizable outputs

### Why Use SuuSri?

- ⏱️ **Save Hours** - Generate presentations in seconds
- 🎨 **Professional Quality** - AI-crafted content
- 💡 **Smart Content** - Relevant, organized information
- 🔧 **Easy to Use** - No technical knowledge needed
- 📱 **Accessible** - Works on all devices
- ♻️ **Reusable** - Download and edit presentations
- 🚀 **Always Available** - 24/7 online access

---

## 🌐 Live Demo

🔗 **[https://ppt-generator-seven-mu.vercel.app](https://ppt-generator-seven-mu.vercel.app)**

Experience the power of AI presentation generation live! Try generating presentations on any topic.

---

## ✨ Features

### 🤖 AI-Powered Generation
- **Intelligent Content Creation** - AI generates relevant, well-structured content
- **Multiple Slide Formats** - Title slides, content slides, conclusion slides
- **Smart Organization** - Logical flow and structure
- **Customizable Prompts** - Create presentations on any topic
- **Multiple AI Models** - Support for Groq API, Gemini API

### 📊 Presentation Management
- **Quick Generation** - Generate presentations in seconds
- **Batch Processing** - Create multiple presentations
- **Template Support** - Pre-defined slide templates
- **Custom Styling** - Professional themes and colors
- **Multi-language Support** - Ready for localization

### 📥 Download & Export
- **PowerPoint Export** - Download as .pptx format
- **Cross-Platform** - Works with PowerPoint, Google Slides, LibreOffice
- **Editable Formats** - Edit presentations after download
- **High-Quality Output** - Professional-grade presentations
- **Instant Download** - No delays, direct downloads

### 🎨 User Interface
- **Responsive Design** - Works on desktop, tablet, mobile
- **Intuitive Interface** - Easy-to-use for all users
- **Real-time Feedback** - Progress indicators during generation
- **Dark/Light Mode** - Comfortable viewing options
- **Accessibility** - WCAG compliant interface

### 🔐 Security
- **Environment Variables** - Secure API key handling
- **Rate Limiting** - Prevent API abuse
- **Input Validation** - Safe input processing
- **Error Handling** - Graceful error messages
- **Privacy** - No data storage between sessions

---

## 🛠️ Tech Stack

### Frontend (44.4% HTML + CSS + JavaScript)

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Bootstrap](https://img.shields.io/badge/Bootstrap-563D7C?style=for-the-badge&logo=bootstrap&logoColor=white)

- **HTML5** (30%) - Semantic markup and structure
- **CSS3** (8.9%) - Styling and responsive design
- **JavaScript** (14.4%) - Interactivity and dynamic features
- **Bootstrap** - Responsive grid system
- **Fetch API** - Asynchronous HTTP requests

### Backend (46.6% Python)

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white)

- **Python 3.8+** - Core backend language
- **Flask** - Lightweight web framework
- **python-pptx** - PowerPoint file generation
- **Requests** - HTTP client for API calls

### AI & LLM Integration

![Groq](https://img.shields.io/badge/Groq%20API-FF6B6B?style=for-the-badge)
![Google](https://img.shields.io/badge/Google%20Gemini-4285F4?style=for-the-badge&logo=google&logoColor=white)

- **Groq API** - High-speed LLM inference
- **Google Gemini API** - Alternative AI model
- **LLM Prompting** - Intelligent prompt engineering

### Deployment

![Vercel](https://img.shields.io/badge/Vercel-000000?style=for-the-badge&logo=vercel&logoColor=white)
![Docker](https://img.shields.io/badge/Docker-2496ED?style=for-the-badge&logo=docker&logoColor=white)

- **Vercel** - Serverless deployment platform
- **Docker** - Containerization (optional)
- **Environment Variables** - Secure configuration

### Configuration Management
- **Procfile** (0.1%) - Heroku/Vercel process definition

---

## 📦 Installation

### Prerequisites
```
Python 3.8 or higher
pip (Python package manager)
Virtual Environment (recommended)
Modern web browser
```

### Step-by-Step Setup

#### 1. Clone the Repository
```bash
git clone https://github.com/SubhamKhandual007/ppt-generator.git
cd ppt-generator
```

#### 2. Create Virtual Environment
```bash
# On Windows
python -m venv venv
venv\Scripts\activate

# On macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

#### 4. Set Up Environment Variables
```bash
# Create .env file
cp .env.example .env

# Edit .env with your API keys
```

#### 5. Run the Application
```bash
# Development
python app.py

# Production
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

#### 6. Access the Application
```
Browser: http://localhost:5000
```

---

## ⚙️ Configuration

### Environment Variables

Create a `.env` file in the project root with the following variables:

```env
# AI API Keys (choose one or both)
GROQ_API_KEY=your_groq_api_key_here
GEMINI_API_KEY=your_google_gemini_api_key_here

# Flask Configuration
FLASK_ENV=production
FLASK_APP=app.py
SECRET_KEY=your_secret_key_here

# API Rate Limiting
RATE_LIMIT=10
RATE_LIMIT_PERIOD=3600

# Presentation Settings
MAX_SLIDES=20
PRESENTATION_WIDTH=960
PRESENTATION_HEIGHT=720
```

### Obtaining API Keys

**Groq API:**
1. Visit [console.groq.com](https://console.groq.com)
2. Create an account
3. Generate API key
4. Add to `.env` file

**Google Gemini API:**
1. Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create API key
3. Add to `.env` file

---

## 🚀 Usage

### Web Interface

1. **Open Application**
   - Navigate to `http://localhost:5000`
   - Or visit the live demo

2. **Enter Topic/Prompt**
   - Type your presentation topic
   - Or paste a detailed prompt
   - Examples:
     - "Climate Change and Global Warming"
     - "Introduction to Machine Learning"
     - "Digital Marketing Strategies for 2024"

3. **Configure Options** (Optional)
   - Select number of slides
   - Choose style/theme
   - Add custom instructions

4. **Generate Presentation**
   - Click "Generate PPT" button
   - Wait for AI to create slides (~5-30 seconds)
   - See progress indicator

5. **Download Presentation**
   - Click "Download" button
   - Presentation saved as `.pptx` file
   - Ready to edit in PowerPoint/Google Slides

### Using the API

#### Generate Presentation

```bash
curl -X POST http://localhost:5000/api/generate \
  -H "Content-Type: application/json" \
  -d '{
    "topic": "Artificial Intelligence",
    "num_slides": 10,
    "style": "professional"
  }'
```

#### Python Example

```python
import requests
import json

url = "http://localhost:5000/api/generate"
payload = {
    "topic": "Web Development Trends",
    "num_slides": 8,
    "style": "creative"
}

response = requests.post(url, json=payload)
presentation_file = response.content

# Save to file
with open("presentation.pptx", "wb") as f:
    f.write(presentation_file)

print("Presentation generated successfully!")
```

#### JavaScript Example

```javascript
const topic = "Machine Learning Basics";
const numSlides = 10;

fetch('/api/generate', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
    },
    body: JSON.stringify({
        topic: topic,
        num_slides: numSlides,
        style: 'professional'
    })
})
.then(response => response.blob())
.then(blob => {
    // Create download link
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'presentation.pptx';
    a.click();
})
.catch(error => console.error('Error:', error));
```

---

## 📁 Project Structure

```
ppt-generator/
│
├── app.py                          # Main Flask application
├── requirements.txt                # Python dependencies
├── Procfile                        # Heroku/Vercel deployment config
├── runtime.txt                     # Python version
├── vercel.json                     # Vercel configuration
├── .env.example                    # Environment variables template
├── .gitignore                      # Git ignore rules
│
├── static/                         # Frontend assets
│   ├── css/
│   │   ├── style.css              # Main stylesheet
│   │   ├── responsive.css         # Responsive design
│   │   └── theme.css              # Theme variables
│   │
│   ├── js/
│   │   ├── main.js                # Main JavaScript
│   │   ├── api.js                 # API communication
│   │   ├── ui.js                  # UI interactions
│   │   └── theme.js               # Theme management
│   │
│   └── images/
│       ├── logo.png
│       ├── favicon.ico
│       └── icons/
│
├── templates/                      # HTML templates
│   ├── base.html                  # Base template
│   ├── index.html                 # Home page
│   ├── generator.html             # Generator interface
│   ├── result.html                # Result display
│   ├── faq.html                   # FAQ page
│   ├── about.html                 # About page
│   └── error.html                 # Error page
│
├── models/                         # Data models (if needed)
│   └── presentation.py
│
├── utils/                          # Utility functions
│   ├── llm_provider.py            # LLM API handlers
│   ├── ppt_generator.py           # PowerPoint generation
│   ├── validators.py              # Input validation
│   └── config.py                  # Configuration management
│
├── tests/                          # Test files
│   ├── test_app.py
│   ├── test_generator.py
│   └── test_api.py
│
└── README.md                       # This file
```

---

## 📡 API Documentation

### Endpoints

#### 1. Health Check
```
GET /api/health
```

**Response:**
```json
{
    "status": "ok",
    "service": "PPT Generator",
    "version": "1.0.0"
}
```

#### 2. Generate Presentation
```
POST /api/generate
Content-Type: application/json
```

**Request Body:**
```json
{
    "topic": "Your Presentation Topic",
    "num_slides": 10,
    "style": "professional",
    "ai_model": "groq",
    "language": "en"
}
```

**Parameters:**
- `topic` (required): Presentation topic or detailed prompt
- `num_slides` (optional): Number of slides (default: 10, max: 20)
- `style` (optional): Presentation style - "professional", "creative", "minimal" (default: "professional")
- `ai_model` (optional): AI model to use - "groq", "gemini" (default: "groq")
- `language` (optional): Content language (default: "en")

**Response:**
```
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation
(Binary PowerPoint file)
```

**Error Response:**
```json
{
    "error": "Error message",
    "status": "error",
    "code": 400
}
```

#### 3. Get Presentation History
```
GET /api/history
```

**Response:**
```json
{
    "presentations": [
        {
            "id": "123abc",
            "topic": "AI Trends",
            "created_at": "2024-05-20T10:30:00Z",
            "num_slides": 10,
            "status": "completed"
        }
    ]
}
```

#### 4. Get Presentation Status
```
GET /api/status/<presentation_id>
```

**Response:**
```json
{
    "id": "123abc",
    "status": "completed",
    "progress": 100,
    "estimated_time_remaining": 0
}
```

---

## 🚀 Deployment

### Deploy to Vercel (Recommended)

1. **Push to GitHub**
```bash
git push origin main
```

2. **Connect to Vercel**
   - Go to [vercel.com](https://vercel.com)
   - Import repository
   - Add environment variables
   - Deploy

3. **Verify Deployment**
   - Application automatically deploys
   - View logs in Vercel dashboard

### Deploy to Heroku

1. **Install Heroku CLI**
```bash
npm install -g heroku
```

2. **Login to Heroku**
```bash
heroku login
```

3. **Create Heroku App**
```bash
heroku create ppt-generator
```

4. **Set Environment Variables**
```bash
heroku config:set GROQ_API_KEY=your_key_here
heroku config:set GEMINI_API_KEY=your_key_here
```

5. **Deploy**
```bash
git push heroku main
```

### Docker Deployment

1. **Build Docker Image**
```bash
docker build -t ppt-generator .
```

2. **Run Container**
```bash
docker run -p 5000:5000 \
  -e GROQ_API_KEY=your_key \
  -e GEMINI_API_KEY=your_key \
  ppt-generator
```

3. **Using Docker Compose**
```bash
docker-compose up -d
```

### Local Server Deployment

```bash
# Production server with Gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app

# With Nginx as reverse proxy
# (Configure Nginx separately)
```

---

## 🔧 Development

### Development Server
```bash
flask run --reload
# or
python app.py
```

### Available Commands
```bash
# Run tests
python -m pytest

# Lint code
pylint app.py
flake8 .

# Format code
black .

# Type checking
mypy app.py
```

### Project Standards
- **Code Style**: PEP 8
- **Testing**: Pytest
- **Documentation**: Docstrings on all functions
- **Git**: Feature branches, meaningful commits

---

## 🐛 Troubleshooting

### Common Issues

**API Key Issues**
```bash
# Verify API key is set
echo $GROQ_API_KEY

# Check .env file exists
ls -la .env

# Update .env with correct key
```

**Module Import Errors**
```bash
# Reinstall dependencies
pip install -r requirements.txt --force-reinstall

# Check Python version
python --version
```

**Port Already in Use**
```bash
# Use different port
PORT=5001 python app.py

# Or kill process on port 5000
lsof -i :5000
kill -9 <PID>
```

**Timeout Issues**
```bash
# Increase timeout in Vercel
# vercel.json: "functions": {"api/**": {"maxDuration": 60}}

# For local: increase timeout
# TIMEOUT=120 python app.py
```

**PowerPoint Generation Issues**
```bash
# Check python-pptx version
pip install python-pptx --upgrade

# Verify file permissions
chmod 755 .

# Check disk space
df -h
```

---

## 📊 Performance Tips

1. **Optimize API Calls**
   - Use caching for repeated topics
   - Batch multiple requests
   - Choose faster AI model (Groq recommended)

2. **Frontend Optimization**
   - Minify CSS/JavaScript
   - Lazy load images
   - Use compression

3. **Backend Optimization**
   - Use connection pooling
   - Cache API responses
   - Implement rate limiting

4. **Deployment**
   - Use CDN for static files
   - Enable gzip compression
   - Optimize database queries

---

## 🤝 Contributing

Contributions are welcome! Here's how to contribute:

1. **Fork the repository**
```bash
git clone https://github.com/yourusername/ppt-generator.git
```

2. **Create feature branch**
```bash
git checkout -b feature/your-feature-name
```

3. **Make improvements**
   - Add new AI models
   - Improve UI/UX
   - Add new presentation styles
   - Improve documentation
   - Add tests

4. **Commit & Push**
```bash
git commit -m "Add: your feature description"
git push origin feature/your-feature-name
```

5. **Open Pull Request**

### Contribution Areas
- 🤖 New AI model integrations
- 🎨 Additional presentation themes
- 📱 Mobile app development
- 🧪 More comprehensive tests
- 📚 Better documentation
- 🌍 Multi-language support
- 🔍 SEO optimization

---

## 📚 Resources

- [Flask Documentation](https://flask.palletsprojects.com)
- [python-pptx Documentation](https://python-pptx.readthedocs.io)
- [Groq API Docs](https://console.groq.com/docs)
- [Google Gemini API](https://ai.google.dev)
- [Vercel Deployment Guide](https://vercel.com/docs)
- [PowerPoint XML Reference](https://learn.microsoft.com/en-us/office/open-xml/overview)

---

## 📞 Support

- 📧 **Email**: subhamkhandual215@gmail.com
- 🐛 **Issues**: [GitHub Issues](https://github.com/SubhamKhandual007/ppt-generator/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/SubhamKhandual007/ppt-generator/discussions)
- 🌐 **Live Demo**: [https://ppt-generator-seven-mu.vercel.app](https://ppt-generator-seven-mu.vercel.app)

---

## 🎓 Author

**Subham Khandual**

B.Tech Computer Science Student | Full Stack Developer | AI/ML Enthusiast

- 🔗 [GitHub](https://github.com/SubhamKhandual007)
- 💼 [LinkedIn](https://www.linkedin.com/in/subham-khandual)
- 📧 [Email](mailto:subhamkhandual215@gmail.com)
- 🌐 [Portfolio](https://portfolio-nine-alpha-8nkzp7nnk6.vercel.app)

---

## 📄 License

This project is licensed under the MIT License - see LICENSE file for details.

---

## 🙏 Acknowledgments

- **Groq** - Fast LLM inference
- **Google** - Gemini AI capabilities
- **Flask** - Web framework
- **python-pptx** - PowerPoint generation library
- **Vercel** - Deployment platform
- **Open Source Community** - Inspiration and support

---

## ⭐ Show Your Support

If you find this project helpful:

- ⭐ **Star the repository**
- 🔗 **Share with others**
- 💬 **Provide feedback**
- 🤝 **Contribute improvements**
- 📢 **Spread the word**

---

<div align="center">

### 🎯 Transform Your Ideas Into Presentations Instantly ✨

**SuuSri AI PPT Generator**

*Where AI Meets Productivity*

Powered by advanced AI | Built with Python | Deployed on Vercel

[Try Live Demo](https://ppt-generator-seven-mu.vercel.app) • [Report Issue](https://github.com/SubhamKhandual007/ppt-generator/issues) • [Request Feature](https://github.com/SubhamKhandual007/ppt-generator/issues)

---

**Last Updated**: May 2026 | **Version**: 1.0.0 | **Status**: Active Development 🚀

[⬆ Back to Top](#-suusri-ai-ppt-generator)

</div>
