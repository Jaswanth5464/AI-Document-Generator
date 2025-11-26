

---


<div align="center">

# 🚀 AI-Powered Document Generator

**Transform your ideas into professional documents and presentations using AI**

<p>
  <a href="https://ai-document-generator-taupe.vercel.app/">
    <img src="https://img.shields.io/badge/🌐_Live_Demo-Visit_App-4CAF50?style=for-the-badge" alt="Live Demo">
  </a>
  <a href="#">
    <img src="https://img.shields.io/badge/📹_Demo-Watch_Video-FF0000?style=for-the-badge" alt="Video Demo">
  </a>
  <a href="#-quick-start">
    <img src="https://img.shields.io/badge/📖_Docs-Read_More-2196F3?style=for-the-badge" alt="Documentation">
  </a>
</p>

<p>
  <img src="https://img.shields.io/badge/React-18.2-61DAFB?logo=react&logoColor=white" alt="React">
  <img src="https://img.shields.io/badge/FastAPI-0.104-009688?logo=fastapi&logoColor=white" alt="FastAPI">
  <img src="https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/Firebase-10.7-FFCA28?logo=firebase&logoColor=black" alt="Firebase">
  <img src="https://img.shields.io/badge/Gemini_AI-Latest-4285F4?logo=google&logoColor=white" alt="Gemini AI">
</p>

---

</div>
## 📋 live link:
https://ai-document-generator-taupe.vercel.app



## 📋 Overview

An intelligent full-stack web application that leverages Google's Gemini AI to generate, refine, and export professional Microsoft Word documents and PowerPoint presentations.

### ✨ Key Features

- 🤖 **AI Content Generation** - Create professional content in seconds
- ✏️ **Smart Refinement** - Improve content with natural language instructions
- 📥 **Export to Office** - Download as .docx or .pptx files
- 💾 **Project Management** - Save and manage multiple documents
- 👍 **Feedback System** - Like/dislike sections and add comments
- 🔐 **Secure Auth** - Firebase authentication

---

## 🛠 Tech Stack

**Frontend:** React, Vite, Firebase, React Router  
**Backend:** FastAPI, Python, Google Gemini AI  
**Database:** Firebase Firestore  
**Hosting:** Vercel (Frontend), Render (Backend)

---

## 🚀 Quick Start

### Prerequisites

- Node.js 18+
- Python 3.11+
- [Google Gemini API Key](https://makersuite.google.com/app/apikey)
- [Firebase Project](https://console.firebase.google.com/)



### Installation

```bash
# Clone repository
git clone https://github.com/YOUR_USERNAME/ai-document-generator.git
cd ai-document-generator

# Backend setup
cd backend
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # Mac/Linux
pip install -r requirements.txt

# Create .env file
echo GEMINI_API_KEY=your_key_here > .env

# Frontend setup
cd ../frontend
npm install
```

### Configuration

**Backend `.env`:**
```env
GEMINI_API_KEY=your_gemini_api_key
```

**Frontend `src/firebase.js`:**
```javascript
const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_AUTH_DOMAIN",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_STORAGE_BUCKET",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};
```

### Run Locally

**Terminal 1 - Backend:**
```bash
cd backend
python main.py
# Runs on http://localhost:8000
```

**Terminal 2 - Frontend:**
```bash
cd frontend
npm run dev
# Runs on http://localhost:5173
```

---

## 📖 Usage

1. **Sign Up/Login** → Create account
2. **Create Project** → Choose Word or PowerPoint
3. **Configure** → Add sections and topic
4. **Generate** → AI creates content
5. **Refine** → Improve with instructions like "make it shorter"
6. **Export** → Download your document

---

## 🌐 Deployment

### Deploy Backend (Render)

1. Push code to GitHub
2. Go to [Render Dashboard](https://dashboard.render.com)
3. New Web Service → Connect GitHub repo
4. Configure:
   - **Root Directory:** `backend`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Add environment variable: `GEMINI_API_KEY`
6. Deploy!

### Deploy Frontend (Vercel)

1. Go to [Vercel Dashboard](https://vercel.com/dashboard)
2. Import Project → Select GitHub repo
3. Configure:
   - **Root Directory:** `frontend`
   - **Build Command:** `npm run build`
   - **Output Directory:** `dist`
4. Update `frontend/src/config.js` with backend URL
5. Deploy!

---

### 📊 System Architecture Flowchart

   <img width="1024" height="1536" alt="image" src="https://github.com/user-attachments/assets/0b948756-9b56-485d-a384-543656a0e036" />

## 📚 API Documentation

### Endpoints

```
POST /api/generate-section   - Generate content for a section
POST /api/refine-section      - Refine existing content
POST /api/export-document     - Export as .docx or .pptx
```

Interactive docs: `http://localhost:8000/docs`

---

## 📁 Project Structure

```
my-document-app/
├── backend/
│   ├── main.py                 # FastAPI application
│   ├── requirements.txt        # Python dependencies
│   ├── .env                    # Environment variables
│   └── utils/
│       ├── docx_generator.py   # Word document generator
│       └── pptx_generator.py   # PowerPoint generator
└── frontend/
    ├── src/
    │   ├── pages/              # React pages
    │   ├── components/         # React components
    │   ├── firebase.js         # Firebase config
    │   └── config.js           # API configuration
    └── package.json
```

---

## 🎥 Demo

**Live Application:** [https://ai-document-generator-taupe.vercel.app/](https://ai-document-generator-taupe.vercel.app/)

**Demo Video:** [Add your video link]

---

## 🐛 Troubleshooting

**Backend won't start:**
```bash
pip install -r requirements.txt
python main.py
```

**Frontend can't connect:**
- Check `config.js` has correct backend URL
- Verify backend is running
- Check browser console for errors

**Export not working:**
- Verify backend logs for errors
- Check if `python-docx` and `python-pptx` installed

---

## 📝 License

MIT License - feel free to use for any purpose

---

## 👤 Author

**Your Name**
- GitHub: [@yourusername](https://github.com/jaswanth5464)
- LinkedIn: [Your LinkedIn](www.linkedin.com/in/jaswanth-kanamrlapudi-a41197252)
- Email: jaswanth5464@gmail.com

---

## 🙏 Acknowledgments

- Google Gemini AI for content generation
- FastAPI for the amazing framework
- Firebase for authentication and database
- React team for the UI library



## ⭐ Show Your Support

Give a ⭐️ if this project helped you!



**Built with ❤️ for the AI Document Generation Assignment**

